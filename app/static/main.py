"""
* You can use pyscript.fetch, but often, you'll run into CORS issues (GitHub is fine though)
* No support for TCP/IP, i.e., no connections with external databases like Postgres
* No access to local file system, but there's a virtual file system where files can be created via URLs or via upload
* TODO: Pictures/Matplotlib should be possible to pass to JS via file system or if not via base64 encoding
* TODO: Look into https://docs.pyscript.net/2024.5.2/user-guide/workers/
"""

import json
import os

os.environ["XLWINGS_LICENSE_KEY"] = "noncommercial"
import xlwings as xw  # noqa: E402
from pyscript import window  # noqa: E402

xwjs = window.xlwings
process_result = window.processResult


async def test(event):
    """Called from task pane button"""
    # Instantiate Book hack
    data = await xwjs.runPython()
    book = xw.Book(json=json.loads(data))

    # Usual xlwings API
    print(book.sheets[0]["A1:A2"].value)
    book.sheets[0]["A3"].value = "xxxxxxx"

    # Process actions (this could be improved so methods are applied immediately)
    process_result(json.dumps(book.json()))


async def hello(name):
    """Used as custom function (a.k.a. UDF)"""
    return [[f"hello from Python, {name}!"]]


window.hello = hello
