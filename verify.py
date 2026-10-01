"""Verify the portable package and run synthetic offline regression tests."""
import hashlib
import importlib.metadata
import json
from pathlib import Path
import subprocess
import sys
import unittest

from playwright.sync_api import sync_playwright
import openpyxl


def main():
    root = Path(__file__).resolve().parent
    manifest = json.loads((root / 'checksums.sha256.json').read_text(encoding='utf-8'))
    for relative, expected in manifest.items():
        path = (root / relative).resolve()
        if not path.is_relative_to(root):
            raise RuntimeError('Unsafe checksum path')
        if hashlib.sha256(path.read_bytes()).hexdigest() != expected:
            raise RuntimeError('File checksum mismatch: ' + relative)
    print(f'PASS: {len(manifest)} packaged files verified', flush=True)
    print('Python', sys.version.split()[0], '| Playwright', importlib.metadata.version('playwright'), '| openpyxl', openpyxl.__version__, flush=True)
    with sync_playwright() as playwright:
        assert playwright.chromium.name == 'chromium'
    print('PASS: bundled Playwright driver started and stopped; no browser opened', flush=True)
    suite = unittest.defaultTestLoader.discover(str(root / 'tests'), pattern='test_*.py')
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        return 1
    # Keep same-named root/skill test modules in separate Python processes.
    skill_tests = root / 'skills/customer-express-issue-query-playwright/tests'
    command = subprocess.run([sys.executable, '-B', '-m', 'unittest', 'discover',
                              '-s', str(skill_tests), '-p', 'test_*.py', '-v'],
                             cwd=root, check=False)
    if command.returncode:
        return command.returncode
    node = root / 'skills/customer-express-issue-query-playwright/runtime/python/Lib/site-packages/playwright/driver/node.exe'
    command = subprocess.run([str(node), '--test', str(root / 'tests/test_query.js')], check=False)
    if command.returncode:
        return command.returncode
    print('PASS: offline self-check complete. Live ERP login/query must be checked separately.', flush=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
