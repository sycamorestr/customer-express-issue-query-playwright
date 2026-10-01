# Third-party distribution notes

Original licenses and notices are retained in the runtime, wheel metadata and Playwright driver distribution. This notice does not replace those licenses.

| Component | Version | Source / license reference |
|---|---|---|
| CPython embeddable Windows x64 | 3.13.12 | https://www.python.org/ftp/python/3.13.12/ ; runtime/python/LICENSE.txt |
| Playwright Python and driver | 1.63.0 | https://pypi.org/project/playwright/1.63.0/ ; installed dist-info and driver/package/LICENSE, NOTICE, ThirdPartyNotices.txt |
| greenlet | 3.5.6 | https://pypi.org/project/greenlet/3.5.6/ ; installed dist-info licenses |
| pyee | 13.0.1 | https://pypi.org/project/pyee/13.0.1/ ; installed dist-info licenses |
| typing_extensions | 4.16.0 | https://pypi.org/project/typing-extensions/4.16.0/ ; installed dist-info licenses |
| openpyxl | 3.1.5 | https://pypi.org/project/openpyxl/3.1.5/ ; installed dist-info license |
| et_xmlfile | 2.0.0 | https://pypi.org/project/et-xmlfile/2.0.0/ ; installed dist-info license |

Runtime paths are relative to `skills/customer-express-issue-query-playwright`. Playwright's official wheel includes the Node executable used by its driver; its third-party notices remain included. No separate OpenCLI, extension, system Node installation or Playwright browser distribution is required.

The original portable CPython and Excel dependencies were retained from the user-supplied package after verifying its complete checksum manifest. Wheel hashes were verified against PyPI metadata; URLs, versions and SHA256 values are in python-dependencies.json. CPython archive provenance from the original package is preserved in download-provenance.json. checksums.sha256.json covers the final distributed files.

CPython uses its isolated python313._pth and bundled site-packages. All dependencies are installed from wheels into that private directory; no system PATH or global Python changes are required. The pip-generated playwright.exe console launcher (which embeds the build machine's interpreter path) is omitted; the portable runner uses the Python API, and the CLI remains available through the bundled python.exe -m playwright. Custom scripts, tests and instructions support the customer-service workflow. No customer workbooks, results, browser profiles or credentials are distributed. Users provide their own authorized ERP session. The package is not endorsed by Microsoft, Python or 聚水潭.
