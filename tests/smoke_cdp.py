"""Optional Windows integration check: a dedicated Edge profile and synthetic local ERP.

Run with the bundled Python: python.exe -B tests/smoke_cdp.py --workdir <empty-test-dir>
No real ERP navigation or order requests are made. The created test browser is closed.
This is deliberately excluded from check.cmd, which must not open a browser.
"""
import argparse
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import socket
import subprocess
import sys
import threading
from urllib.parse import parse_qs

import openpyxl
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / 'skills/customer-express-issue-query-playwright/scripts'
NUMBERS = ['JT0000000001', 'YT0000000002', 'ZT0000000003']


class Fixture(BaseHTTPRequestHandler):
    requests = []
    rechecked = False

    def log_message(self, *args):
        pass

    def reply(self, body, content_type):
        raw = body.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        if self.path == '/':
            body = '<title>Synthetic CDP test</title><iframe src="/shell"></iframe>'
        elif self.path == '/shell':
            body = '<iframe src="/app/order/order/list.aspx"></iframe>'
        elif self.path == '/app/order/order/list.aspx':
            body = '<title>Synthetic order frame</title><input id="__VIEWSTATE" value="synthetic-viewstate">'
        else:
            self.send_error(404)
            return
        self.reply(body, 'text/html; charset=utf-8')

    def do_POST(self):
        assert self.path.startswith('/app/order/order/list.aspx?')
        form = parse_qs(self.rfile.read(int(self.headers['Content-Length'])).decode())
        assert form['__VIEWSTATE'] == ['synthetic-viewstate']
        callback = json.loads(form['__CALLBACKPARAM'][0])
        assert callback['Method'] == 'LoadDataToJSON'
        filters = json.loads(callback['Args'][1])
        assert len(filters) == 1 and filters[0]['k'] == 'l_id' and filters[0]['c'] == '@='
        tokens = filters[0]['v'].split(',')
        nos = [n for n in tokens if not n.startswith('@')]
        assert set(tokens) == set(nos + ['@' + n for n in nos])
        self.requests.append(nos)
        rows = []
        for no in nos:
            if no == NUMBERS[1] and not self.rechecked:
                continue
            for index in range(2 if no == NUMBERS[0] else 1):
                rows.append({'o_id': no + str(index), 'l_id': '@' + no,
                             'confirm_date': '2026-01-01 12:00:00', 'paid_amount': 0 if index == 0 else 12,
                             'remark': 'synthetic | note\n complete'})
        body = json.dumps({'IsSuccess': True, 'ReturnValue': json.dumps({'datas': rows})})
        self.reply('0|' + body, 'text/plain; charset=utf-8')


def run_cli(*args, expected=0):
    command = subprocess.run([sys.executable, '-B', *map(str, args)], capture_output=True,
                             encoding='utf-8', timeout=45, creationflags=subprocess.CREATE_NO_WINDOW)
    if command.returncode != expected:
        raise AssertionError(f'CLI failed ({command.returncode}): {command.stdout} {command.stderr}')
    return command.stdout


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workdir', required=True)
    args = parser.parse_args()
    work = Path(args.workdir).resolve()
    work.mkdir(parents=True, exist_ok=True)
    assert not any(work.iterdir()), 'Use an empty test directory'
    with socket.socket() as sock:
        sock.bind(('127.0.0.1', 0))
        cdp_port = sock.getsockname()[1]
    server = ThreadingHTTPServer(('127.0.0.1', 0), Fixture)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    endpoint = f'http://127.0.0.1:{cdp_port}'
    start = ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(SCRIPTS / 'start_browser.ps1'),
             '-Browser', 'Edge', '-Port', str(cdp_port), '-UserDataDir', str(work / 'browser-profile')]
    started = subprocess.run(start, capture_output=True, encoding='utf-8', errors='replace', timeout=45,
                             creationflags=subprocess.CREATE_NO_WINDOW)
    assert started.returncode == 0, started.stdout + started.stderr
    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.connect_over_cdp(endpoint)
            try:
                doctor_script = SCRIPTS / 'cdp_query.py'
                before = json.loads(run_cli(doctor_script, 'doctor', '--endpoint', endpoint, expected=2))
                assert before['connected'] and not before['ok']
                page = browser.contexts[0].new_page()
                page.goto(f'http://127.0.0.1:{server.server_port}/', wait_until='load')
                doctor = json.loads(run_cli(doctor_script, 'doctor', '--endpoint', endpoint, '--allowed-host', '127.0.0.1'))
                assert doctor['ok'] and len(doctor['orderFrames']) == 1
                reused = subprocess.run(start, capture_output=True, encoding='utf-8', errors='replace', timeout=45,
                                        creationflags=subprocess.CREATE_NO_WINDOW)
                assert reused.returncode == 0 and 'Reusing' in reused.stdout

                source = work / 'synthetic.xlsx'
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(['快递单号', '原字段'])
                for no in NUMBERS:
                    ws.append([no, '=1+2'])
                wb.save(source)
                job = work / 'job'
                run_cli(SCRIPTS / 'excel.py', 'prepare', '--input', source, '--workdir', job)
                query = [doctor_script, 'query', '--endpoint', endpoint, '--allowed-host', '127.0.0.1',
                         '--workdir', job, '--batch-size', '2']
                run_cli(*query, '--max-batches', '1')
                results = json.loads((job / 'results.json').read_text('utf-8'))
                assert list(results) == NUMBERS[:2] and results[NUMBERS[0]]['matched'] == 2
                assert results[NUMBERS[0]]['data'][0][2] == '0'
                assert results[NUMBERS[0]]['data'][0][3] == 'synthetic | note complete'
                assert results[NUMBERS[1]]['matched'] == 0
                assert browser.is_connected() and not page.is_closed(), 'Runner must leave the owned browser open'
                run_cli(*query)
                assert Fixture.requests == [NUMBERS[:2], NUMBERS[2:]]
                Fixture.rechecked = True
                run_cli(*query, '--recheck-missing')
                assert Fixture.requests[-1] == [NUMBERS[1]]
                result = json.loads(run_cli(SCRIPTS / 'excel.py', 'export', '--workdir', job))
                assert result['matched'] == 3 and result['notFound'] == 0 and result['sourceUnchanged']
                out = openpyxl.load_workbook(result['output'])
                assert out.active['B2'].value == '=1+2'
                out.close()
                (work / 'smoke-summary.json').write_text(json.dumps({
                    'ok': True, 'realCdp': True, 'browser': 'Edge', 'nestedFrames': True,
                    'launchReuse': True, 'resume': True, 'recheckMissing': True, 'zeroAmount': True,
                    'exportSourceUnchanged': True, 'realErpQueried': False,
                    'syntheticRequestBatches': len(Fixture.requests)}, indent=2), encoding='utf-8')
                print('PASS: real Edge CDP, launcher reuse, nested frames, asynchronous query, checkpoint/resume, recheck, Excel export, browser preserved')
            finally:
                # This test created and owns the dedicated browser/profile; never use this on a user's browser.
                browser.new_browser_cdp_session().send('Browser.close')
    finally:
        server.shutdown()
        server.server_close()


if __name__ == '__main__':
    main()
