"""
* You can use pyscript.fetch, but often, you'll run into CORS issues (GitHub is fine though)
* No support for TCP/IP, i.e., no connections with external databases like Postgres
* No access to local file system, but there's a virtual file system where files can be created via URLs or via upload
* Pictures/Matplotlib should be possible to pass to JS via file system or if not via base64 encoding
* TODO: test out custom functions
"""

import json
import os
import sys

os.environ["XLWINGS_LICENSE_KEY"] = "noncommercial"
import xlwings as xw  # noqa: E402
from pyscript import window  # noqa: E402


async def test(event):
    print(xw.__version__)
    print(sys.version)

    # xlwings.js has the version that is included in base.html
    xwjs = window.xlwings
    print(await xwjs.getActiveBookName())

    # Or don't return JSON.stringify in runPython and do data.to_py() instead
    data = await xwjs.runPython()
    data = json.loads(data)

    book = xw.Book(json=data)
    print(book.sheets[0]["A1:A2"].value)
    book.sheets[0]["A3"].value = "xxxxxxx"

    # Result
    process_result = window.processResult
    process_result(json.dumps(book.json()))
