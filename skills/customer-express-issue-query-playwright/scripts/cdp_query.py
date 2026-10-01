"""Attach to a user's CDP browser and query the high-edition ERP order frame."""
import argparse
import contextlib
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import tempfile
import time
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright, Error as PlaywrightError

ORDER_PATH = '/app/order/order/list.aspx'
SCRIPT_DIR = Path(__file__).resolve().parent
SAFE_REASONS = frozenset((
    'Invalid tracking batch', 'Invalid request timeout', 'Open the high edition order page first',
    'Order request timed out; check the browser and retry',
    'Order request failed; check browser login, verification, or connection',
    'Invalid response; check login or verification in browser',
    'API unsuccessful; inspect browser permissions or verification',
    'Unexpected order response schema',
    'Single tracking number reaches page limit; cannot certify completeness',
    'Order query failed; inspect the browser and retry',
))


class QueryError(Exception):
    pass


def require(condition, message):
    if not condition:
        raise QueryError(message)


def load_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def endpoint_url(value):
    try:
        parsed = urlsplit(value)
        valid = (parsed.scheme == 'http' and parsed.hostname in ('127.0.0.1', 'localhost', '::1')
                 and parsed.port is not None and not parsed.username and not parsed.password
                 and parsed.path in ('', '/') and not parsed.query and not parsed.fragment)
    except ValueError:
        valid = False
    if not valid:
        raise argparse.ArgumentTypeError('Use a loopback HTTP CDP endpoint, e.g. http://127.0.0.1:9222')
    return value.rstrip('/')


def positive_int(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError('Must be positive')
    return number


def trusted_frame(url, allowed_host=None):
    p = urlsplit(url)
    host = (p.hostname or '').lower()
    trusted = host == 'erp321.com' or host.endswith('.erp321.com')
    if allowed_host:
        trusted = trusted or host == allowed_host.lower()
    scheme_ok = p.scheme == 'https' or (p.scheme == 'http' and host in ('127.0.0.1', 'localhost', '::1'))
    return trusted and scheme_ok and p.path.lower() == ORDER_PATH


def safe_url(url):
    p = urlsplit(url)
    return f'{p.scheme}://{p.netloc.split("@")[-1]}{p.path}'


def find_frames(browser, page_index=None, allowed_host=None):
    pages = [p for c in browser.contexts for p in c.pages if not p.is_closed()]
    require(page_index is None or 1 <= page_index <= len(pages), 'Page index no longer exists; run doctor again')
    found = []
    for index, page in enumerate(pages, 1):
        if page_index is not None and index != page_index:
            continue
        for frame in page.frames:
            if trusted_frame(frame.url, allowed_host):
                try:
                    ready = frame.evaluate("() => Boolean(document.getElementById('__VIEWSTATE'))")
                except PlaywrightError:
                    ready = False
                found.append({'pageIndex': index, 'pageUrl': safe_url(page.url),
                              'frameUrl': safe_url(frame.url), 'ready': ready, 'frame': frame})
    return found


def select_frame(browser, page_index=None, allowed_host=None):
    candidates = find_frames(browser, page_index, allowed_host)
    require(candidates, 'No high-edition order frame found; open and log into the order list, then run doctor')
    require(len(candidates) == 1, 'Multiple order frames found; keep one target order page or choose --page-index from doctor')
    require(candidates[0]['ready'], 'Order page is not ready or login expired; open the order list and run doctor')
    return candidates[0]


@contextlib.contextmanager
def connection(endpoint, timeout_ms):
    # Stopping the local Playwright driver disconnects CDP without closing the user's browser.
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(endpoint, timeout=timeout_ms)
        yield browser


@contextlib.contextmanager
def work_lock(work):
    with (work / 'query.lock').open('a+b') as stream:
        stream.seek(0, 2)
        if stream.tell() == 0:
            stream.write(b'\0')
            stream.flush()
        stream.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise QueryError('Another query is using this work directory') from exc
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == 'nt':
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream, fcntl.LOCK_UN)


def atomic_save(path, value):
    fd, tmp = tempfile.mkstemp(prefix='.results-', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write('\n')
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(tmp, path)
    finally:
        Path(tmp).unlink(missing_ok=True)


def validate_items(items, expected):
    require(isinstance(items, list) and len(items) == len(expected), 'Incomplete batch/checkpoint response')
    checked = {}
    for item in items:
        require(isinstance(item, dict), 'Invalid result item')
        no, count, rows = item.get('no'), item.get('matched'), item.get('data')
        require(isinstance(no, str) and no in expected and no not in checked, 'Unexpected or duplicate tracking number')
        require(type(count) is int and count >= 0 and isinstance(rows, list), 'Invalid match count/data')
        require(len(rows) == max(1, count), 'Invalid result row count')
        for row in rows:
            require(isinstance(row, list) and len(row) == 4 and row[0] == no and all(isinstance(v, str) for v in row), 'Invalid four-field result')
        require(count != 0 or rows == [[no, '', '', '']], 'Invalid not-found result')
        checked[no] = {'no': no, 'matched': count, 'data': rows}
    require(set(checked) == set(expected), 'Incomplete tracking number coverage')
    return checked


def read_job(work):
    manifest = load_json(work / 'manifest.json')
    require(isinstance(manifest, dict) and manifest.get('schemaVersion') == 1, 'Unsupported manifest')
    unique = manifest.get('unique')
    require(isinstance(unique, list) and unique and all(isinstance(n, str) and re.fullmatch(r'[A-Z0-9]{8,}', n) and re.search(r'[0-9]', n) for n in unique), 'Invalid manifest tracking numbers')
    require(len(unique) == len(set(unique)), 'Duplicate manifest tracking numbers')
    mapped = list(dict.fromkeys(no for sheet in manifest['sheets'] for row in sheet['rows'] for no in row['nos']))
    require(mapped == unique, 'Manifest row mapping changed')
    require(digest(manifest['source']) == manifest['sha256'], 'Source workbook changed; use a new work directory')
    path = work / 'results.json'
    saved = load_json(path) if path.exists() else {}
    require(isinstance(saved, dict) and set(saved) <= set(unique), 'Checkpoint contains foreign tracking numbers')
    # Check dictionary keys as well as their embedded no fields.
    require(all(isinstance(v, dict) and v.get('no') == k for k, v in saved.items()), 'Checkpoint key mismatch')
    saved = validate_items(list(saved.values()), list(saved))
    return manifest, saved


def public_logs(logs):
    require(isinstance(logs, list), 'Invalid query log')
    clean = []
    for item in logs:
        require(isinstance(item, dict), 'Invalid query log item')
        entry = {}
        for key in ('requested', 'rows', 'matched', 'fallback'):
            if key in item:
                require(type(item[key]) is int and item[key] >= 0, 'Invalid query count')
                entry[key] = item[key]
        if item.get('split') is True:
            entry['split'] = True
        clean.append(entry)
    return clean


def run_query(args):
    work = Path(args.workdir).resolve(strict=True)
    require(work.is_dir(), 'Work directory missing; run excel.py prepare first')
    with work_lock(work):
        manifest, results = read_job(work)
        manifest_hash = digest(work / 'manifest.json')
        pending = [n for n in manifest['unique'] if n not in results or (args.recheck_missing and results[n]['matched'] == 0)]
        if not pending:
            return {'ok': True, 'completed': len(results), 'total': len(manifest['unique']), 'queriedBatches': 0, 'message': 'No pending tracking numbers'}
        query_js = (SCRIPT_DIR / 'query.js').read_text(encoding='utf-8')
        batches = 0
        with connection(args.endpoint, args.connect_timeout_ms) as browser:
            target = select_frame(browser, args.page_index, args.allowed_host)
            frame = target['frame']
            for start in range(0, len(pending), args.batch_size):
                if args.max_batches and batches >= args.max_batches:
                    break
                require(digest(work / 'manifest.json') == manifest_hash, 'Manifest changed while querying; checkpoint retained')
                require(digest(manifest['source']) == manifest['sha256'], 'Source changed while querying; checkpoint retained')
                current = select_frame(browser, args.page_index, args.allowed_host)
                require(current['frame'] == frame and current['frameUrl'] == target['frameUrl'], 'Order page changed; run doctor and resume')
                batch = pending[start:start + args.batch_size]
                response = frame.evaluate(query_js, {'nos': batch, 'timeoutMs': args.request_timeout_ms})
                require(isinstance(response, dict), 'Invalid browser response; checkpoint retained')
                # query.js returns only fixed diagnostic strings; never expose Playwright exceptions or response bodies.
                if response.get('ok') is not True:
                    reason = response.get('reason')
                    safe = isinstance(reason, str) and (reason in SAFE_REASONS or re.fullmatch(r'HTTP [1-5][0-9]{2}; check browser login or permissions', reason))
                    require(False, 'Query stopped; checkpoint retained. ' + (reason if safe else 'Inspect browser login/verification'))
                checked = validate_items(response.get('results'), batch)
                logs = public_logs(response.get('logs'))
                require(digest(work / 'manifest.json') == manifest_hash and digest(manifest['source']) == manifest['sha256'], 'Input changed during request; current batch discarded')
                results.update(checked)
                atomic_save(work / 'results.json', results)
                with (work / 'query-log.jsonl').open('a', encoding='utf-8', newline='\n') as stream:
                    stream.write(json.dumps({'batch': batches + 1, 'completed': len(results), 'logs': logs}, ensure_ascii=False) + '\n')
                batches += 1
                matched = sum(item['matched'] > 0 for item in results.values())
                print(json.dumps({'completed': len(results), 'total': len(manifest['unique']), 'matched': matched,
                                  'notFound': len(results) - matched, 'fallback': sum(x.get('fallback', 0) for x in logs)}), flush=True)
                if start + args.batch_size < len(pending) and (not args.max_batches or batches < args.max_batches):
                    time.sleep(args.delay_ms / 1000)
        return {'ok': True, 'completed': len(results), 'total': len(manifest['unique']), 'queriedBatches': batches,
                'complete': len(results) == len(manifest['unique'])}


def doctor(args):
    with connection(args.endpoint, args.connect_timeout_ms) as browser:
        candidates = find_frames(browser, args.page_index, args.allowed_host)
        return {'ok': len(candidates) == 1 and candidates[0]['ready'], 'connected': True,
                'browserVersion': browser.version, 'orderFrames': [{k: v for k, v in c.items() if k != 'frame'} for c in candidates],
                'message': 'Verify the correct ERP account in the browser before querying. Select one ready order frame.'}


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    for name in ('doctor', 'query'):
        command = sub.add_parser(name)
        command.add_argument('--endpoint', type=endpoint_url, default='http://127.0.0.1:9222')
        command.add_argument('--page-index', type=positive_int, help='1-based page index reported by doctor')
        command.add_argument('--allowed-host', help='Additional exact trusted order-page hostname; use only a verified ERP host')
        command.add_argument('--connect-timeout-ms', type=positive_int, default=10000)
        if name == 'query':
            command.add_argument('--workdir', required=True)
            command.add_argument('--batch-size', type=int, choices=range(1, 61), default=25, metavar='1..60')
            command.add_argument('--delay-ms', type=int, default=150)
            command.add_argument('--max-batches', type=int, default=0)
            command.add_argument('--recheck-missing', action='store_true')
            command.add_argument('--request-timeout-ms', type=positive_int, default=30000)
    args = parser.parse_args()
    try:
        if args.command == 'query':
            require(args.delay_ms >= 120 and args.max_batches >= 0, 'delay-ms must be >=120 and max-batches >=0')
            require(args.request_timeout_ms <= 300000, 'request-timeout-ms must be <=300000')
        result = doctor(args) if args.command == 'doctor' else run_query(args)
        print(json.dumps(result, ensure_ascii=False), flush=True)
        return 0 if result['ok'] else 2
    except QueryError as exc:
        error = str(exc)
    except PlaywrightError:
        error = 'CDP connection/page operation failed; check endpoint, browser, login and order page. Checkpoint retained.'
    except (OSError, ValueError, KeyError, TypeError):
        error = 'Invalid or inaccessible input/checkpoint/runtime; inspect files and permissions. Checkpoint retained.'
    print(json.dumps({'ok': False, 'error': error}, ensure_ascii=False), file=sys.stderr)
    return 1


if __name__ == '__main__':
    sys.exit(main())
