import sys

import xlwings as xw
from pyscript import window


async def test(x):
    print(xw.__version__)
    print(sys.version)

    # xlwings.js has the version that is included in base.html
    xlwingsjs = window.xlwings

    print(await xlwingsjs.getActiveBookName())
