"""Offline behavior tests for CDP querying; no browser or ERP is contacted."""
import argparse
import contextlib
import copy
import importlib.util
import io
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch


SCRIPT = (Path(__file__).resolve().parents[1] / "skills" /
          "customer-express-issue-query-playwright" / "scripts" / "cdp_query.py")
SPEC = importlib.util.spec_from_file_location("cdp_runner", SCRIPT)
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)
ORDER_URL = "https://www.erp321.com/app/order/order/list.aspx"


def item(no, count=1):
    rows = [[no, "2026-01-01 12:00:00", "0", "synthetic"]] if count else [[no, "", "", ""]]
    return {"no": no, "matched": count, "data": rows}


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")


class FakeFrame:
    def __init__(self, url=ORDER_URL, responder=None, ready=True):
        self.url = url
        self.ready = ready
        self.batches = []
        self.responder = responder

    def evaluate(self, script, arg=None):
        if arg is None:
            return self.ready
        self.batches.append(list(arg["nos"]))
        if self.responder:
            return self.responder(arg["nos"])
        return {"ok": True, "results": [item(no) for no in arg["nos"]],
                "logs": [{"requested": len(arg["nos"]), "matched": len(arg["nos"])}]}


def fake_page(*frames, url="https://www.erp321.com/home?secret=synthetic", closed=False):
    return SimpleNamespace(frames=list(frames), url=url, is_closed=lambda: closed)


def fake_browser(*pages):
    return SimpleNamespace(contexts=[SimpleNamespace(pages=list(pages))], version="synthetic")


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="express-runner-test-")
        self.addCleanup(self.tmp.cleanup)
        self.work = Path(self.tmp.name)
        self.source = self.work / "synthetic.xlsx"
        self.source.write_bytes(b"synthetic source bytes; not customer data")
        self.nos = [f"AA0000000{i}" for i in range(1, 6)]
        self.manifest = {
            "schemaVersion": 1, "source": str(self.source), "sha256": runner.digest(self.source),
            "unique": self.nos,
            "sheets": [{"name": "Synthetic", "rows": [
                {"row": index + 2, "nos": [no]} for index, no in enumerate(self.nos)
            ]}],
        }
        write_json(self.work / "manifest.json", self.manifest)
        self.args = argparse.Namespace(
            workdir=str(self.work), endpoint="http://127.0.0.1:9222", connect_timeout_ms=1000,
            page_index=None, allowed_host=None, batch_size=2, max_batches=0,
            request_timeout_ms=1000, delay_ms=120, recheck_missing=False,
        )

    def run_fake(self, frame):
        browser = fake_browser(fake_page(frame))
        with patch.object(runner, "connection", return_value=contextlib.nullcontext(browser)), \
                patch.object(runner.time, "sleep"), contextlib.redirect_stdout(io.StringIO()):
            return runner.run_query(self.args)

    def results(self):
        return json.loads((self.work / "results.json").read_text(encoding="utf-8"))

    def invoke_main(self, **overrides):
        stdout, stderr = io.StringIO(), io.StringIO()
        argv = ["cdp_query.py", "query", "--workdir", str(self.work)]
        with patch.object(runner.sys, "argv", argv), contextlib.redirect_stdout(stdout), \
                contextlib.redirect_stderr(stderr), patch.multiple(runner, **overrides):
            exit_code = runner.main()
        return exit_code, stdout.getvalue(), stderr.getvalue()

    def test_batch_validation_rejects_foreign_duplicate_missing_and_boolean_counts(self):
        a, b = self.nos[:2]
        variants = {
            "foreign": [item(a), item("ZZ00000000")],
            "duplicate": [item(a), item(a)],
            "missing": [item(a)],
            "false count": [item(a, False), item(b)],
            "true count": [item(a, True), item(b)],
            "negative count": [dict(item(a), matched=-1), item(b)],
            "wrong row no": [dict(item(a), data=[[b, "", "", ""]]), item(b)],
        }
        for label, items in variants.items():
            with self.subTest(label=label), self.assertRaises(runner.QueryError):
                runner.validate_items(items, [a, b])
        accepted = runner.validate_items([item(a, 0), item(b)], [a, b])
        self.assertEqual(accepted[a]["matched"], 0)
        self.assertEqual(accepted[b]["data"][0][2], "0")

    def test_corrupt_checkpoints_stop_before_any_connection_and_hide_raw_data(self):
        secret = "synthetic-token=do-not-report"
        variants = [
            "{" + secret,
            json.dumps([item(self.nos[0])]),
            json.dumps({"ZZ00000000": item("ZZ00000000")}),
            json.dumps({self.nos[0]: item(self.nos[1])}),
            json.dumps({self.nos[0]: dict(item(self.nos[0]), matched=False)}),
        ]
        for contents in variants:
            with self.subTest(contents=contents[:15]):
                (self.work / "results.json").write_text(contents, encoding="utf-8")
                connect = MagicMock()
                code, stdout, stderr = self.invoke_main(connection=connect)
                self.assertEqual(code, 1)
                connect.assert_not_called()
                self.assertNotIn(secret, stdout + stderr)
                self.assertEqual((self.work / "results.json").read_text(encoding="utf-8"), contents)

    def test_changed_source_or_manifest_is_rejected_before_connection(self):
        bad_mapping = copy.deepcopy(self.manifest)
        bad_mapping["sheets"][0]["rows"][0]["nos"] = [self.nos[1]]
        write_json(self.work / "manifest.json", bad_mapping)
        with patch.object(runner, "connection") as connect, self.assertRaises(runner.QueryError):
            runner.run_query(self.args)
        connect.assert_not_called()
        write_json(self.work / "manifest.json", self.manifest)
        self.source.write_bytes(b"changed source")
        with patch.object(runner, "connection") as connect, self.assertRaises(runner.QueryError):
            runner.run_query(self.args)
        connect.assert_not_called()

    def test_input_changes_during_request_discard_the_whole_batch(self):
        for changed in ("source", "manifest"):
            with self.subTest(changed=changed):
                self.source.write_bytes(b"synthetic source bytes; not customer data")
                write_json(self.work / "manifest.json", self.manifest)
                saved = {self.nos[-1]: item(self.nos[-1])}
                write_json(self.work / "results.json", saved)

                def respond(nos):
                    if changed == "source":
                        self.source.write_bytes(b"modified during request")
                    else:
                        with (self.work / "manifest.json").open("a", encoding="utf-8") as stream:
                            stream.write("\n")
                    return {"ok": True, "results": [item(no) for no in nos], "logs": []}

                with self.assertRaisesRegex(runner.QueryError, "Input changed"):
                    self.run_fake(FakeFrame(responder=respond))
                self.assertEqual(self.results(), saved)
                self.assertFalse((self.work / "query-log.jsonl").exists())

    def test_only_a_complete_valid_batch_is_committed(self):
        batches = 0

        def respond(nos):
            nonlocal batches
            batches += 1
            results = [item(no) for no in nos]
            if batches == 2:
                results.pop()
            return {"ok": True, "results": results, "logs": []}

        with self.assertRaisesRegex(runner.QueryError, "Incomplete"), \
                patch.object(runner, "atomic_save", wraps=runner.atomic_save) as save:
            self.run_fake(FakeFrame(responder=respond))
        self.assertEqual(self.results(), {no: item(no) for no in self.nos[:2]})
        self.assertEqual(save.call_count, 1)
        self.assertEqual(len((self.work / "query-log.jsonl").read_text(encoding="utf-8").splitlines()), 1)

    def test_invalid_logs_cannot_commit_even_when_batch_results_are_valid(self):
        def respond(nos):
            return {"ok": True, "results": [item(no) for no in nos], "logs": [{"matched": False}]}

        with self.assertRaisesRegex(runner.QueryError, "Invalid query count"):
            self.run_fake(FakeFrame(responder=respond))
        self.assertFalse((self.work / "results.json").exists())
        self.assertFalse((self.work / "query-log.jsonl").exists())

    def test_resume_skips_completed_entries_and_rechecks_only_missing(self):
        saved = {self.nos[0]: item(self.nos[0]), self.nos[1]: item(self.nos[1], 0)}
        write_json(self.work / "results.json", saved)
        frame = FakeFrame()
        summary = self.run_fake(frame)
        self.assertEqual(frame.batches, [self.nos[2:4], self.nos[4:]])
        self.assertTrue(summary["complete"])
        self.assertEqual(self.results()[self.nos[1]], saved[self.nos[1]])
        self.args.recheck_missing = True
        recheck = FakeFrame()
        self.run_fake(recheck)
        self.assertEqual(recheck.batches, [[self.nos[1]]])
        self.assertEqual(self.results()[self.nos[1]]["matched"], 1)
        with patch.object(runner, "connection") as connect:
            done = runner.run_query(self.args)
        connect.assert_not_called()
        self.assertEqual(done["queriedBatches"], 0)

    def test_max_batches_stops_after_saving_and_next_run_resumes(self):
        self.args.max_batches = 1
        first = FakeFrame()
        summary = self.run_fake(first)
        self.assertEqual(first.batches, [self.nos[:2]])
        self.assertEqual(summary["queriedBatches"], 1)
        self.assertFalse(summary["complete"])
        self.assertEqual(list(self.results()), self.nos[:2])
        next_run = FakeFrame()
        self.run_fake(next_run)
        self.assertEqual(next_run.batches, [self.nos[2:4]])

    def test_atomic_save_leaves_previous_checkpoint_intact_when_replace_fails(self):
        path = self.work / "results.json"
        old = {self.nos[0]: item(self.nos[0])}
        write_json(path, old)
        before = path.read_bytes()
        with patch.object(runner.os, "replace", side_effect=OSError("synthetic failure")), \
                self.assertRaises(OSError):
            runner.atomic_save(path, {self.nos[1]: item(self.nos[1])})
        self.assertEqual(path.read_bytes(), before)
        self.assertEqual(list(self.work.glob(".results-*.tmp")), [])
        runner.atomic_save(path, {self.nos[1]: item(self.nos[1])})
        self.assertEqual(self.results(), {self.nos[1]: item(self.nos[1])})

    def test_work_lock_rejects_a_second_holder_and_releases_after_exception(self):
        with self.assertRaisesRegex(RuntimeError, "synthetic"):
            with runner.work_lock(self.work):
                with self.assertRaisesRegex(runner.QueryError, "Another query"):
                    with runner.work_lock(self.work):
                        self.fail("Second holder acquired the same work directory")
                raise RuntimeError("synthetic")
        with runner.work_lock(self.work):
            pass

    def test_frame_selection_rejects_ambiguity_expired_login_and_untrusted_hosts(self):
        first, second = FakeFrame(), FakeFrame()
        browser = fake_browser(fake_page(first), fake_page(second))
        with self.assertRaisesRegex(runner.QueryError, "Multiple order frames"):
            runner.select_frame(browser)
        self.assertIs(runner.select_frame(browser, page_index=2)["frame"], second)
        with self.assertRaisesRegex(runner.QueryError, "Multiple order frames"):
            runner.select_frame(fake_browser(fake_page(first, second)), page_index=1)
        with self.assertRaisesRegex(runner.QueryError, "not ready"):
            runner.select_frame(fake_browser(fake_page(FakeFrame(ready=False))))
        for url in (
            "https://erp321.com.attacker.invalid/app/order/order/list.aspx",
            "https://attacker-erp321.com/app/order/order/list.aspx",
            "http://www.erp321.com/app/order/order/list.aspx",
            "https://www.erp321.com/not-an-order-page",
        ):
            with self.subTest(url=url):
                self.assertFalse(runner.trusted_frame(url))
        local = "http://127.0.0.1/app/order/order/list.aspx"
        self.assertFalse(runner.trusted_frame(local))
        self.assertTrue(runner.trusted_frame(local, "127.0.0.1"))
        self.assertTrue(runner.trusted_frame("https://verified.invalid/app/order/order/list.aspx", "verified.invalid"))
        self.assertFalse(runner.trusted_frame("https://sub.verified.invalid/app/order/order/list.aspx", "verified.invalid"))
        self.assertEqual(runner.find_frames(fake_browser(fake_page(first, closed=True))), [])
        with self.assertRaisesRegex(runner.QueryError, "Page index"):
            runner.select_frame(browser, page_index=3)

    def test_endpoint_is_local_and_diagnostics_strip_credentials_and_raw_logs(self):
        self.assertEqual(runner.endpoint_url("http://127.0.0.1:9222/"), "http://127.0.0.1:9222")
        for endpoint in ("http://remote.invalid:9222", "http://user:password@localhost:9222",
                         "http://localhost:9222?token=synthetic", "https://localhost:9222",
                         "http://localhost:9222/path", "http://localhost"):
            with self.subTest(endpoint=endpoint), self.assertRaises(argparse.ArgumentTypeError):
                runner.endpoint_url(endpoint)
        self.assertEqual(runner.safe_url("https://user:password@www.erp321.com/path?token=secret#private"),
                         "https://www.erp321.com/path")
        self.assertEqual(runner.public_logs([{"requested": 2, "rows": 1, "raw": "secret order body",
                                              "token": "secret cookie", "split": True}]),
                         [{"requested": 2, "rows": 1, "split": True}])

    def test_connection_stops_driver_without_closing_browser_on_success_or_error(self):
        for fail_inside in (False, True):
            with self.subTest(fail_inside=fail_inside):
                context = MagicMock()
                playwright = context.__enter__.return_value
                browser = playwright.chromium.connect_over_cdp.return_value
                with patch.object(runner, "sync_playwright", return_value=context):
                    try:
                        with runner.connection("http://127.0.0.1:9222", 1234) as connected:
                            self.assertIs(connected, browser)
                            if fail_inside:
                                raise RuntimeError("synthetic")
                    except RuntimeError:
                        self.assertTrue(fail_inside)
                playwright.chromium.connect_over_cdp.assert_called_once_with("http://127.0.0.1:9222", timeout=1234)
                context.__exit__.assert_called_once()
                browser.close.assert_not_called()

    def test_failed_browser_response_never_echoes_arbitrary_sensitive_reason(self):
        secret = "cookie=synthetic-private; raw order body: customer-data"
        browser = fake_browser(fake_page(FakeFrame(responder=lambda nos: {"ok": False, "reason": secret})))
        connect = MagicMock(return_value=contextlib.nullcontext(browser))
        code, stdout, stderr = self.invoke_main(connection=connect)
        self.assertEqual(code, 1)
        self.assertNotIn(secret, stdout + stderr)
        self.assertFalse((self.work / "results.json").exists())

    def test_playwright_errors_do_not_echo_url_credentials_or_response_bodies(self):
        secret = "ws://user:password@localhost:9222/?token=synthetic-private raw order body"
        connect = MagicMock(side_effect=runner.PlaywrightError(secret))
        code, stdout, stderr = self.invoke_main(connection=connect)
        self.assertEqual(code, 1)
        self.assertNotIn(secret, stdout + stderr)
        self.assertIn("CDP connection/page operation failed", stderr)


if __name__ == "__main__":
    unittest.main()
