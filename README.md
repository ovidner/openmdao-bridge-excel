OpenMDAO Bridge for Excel
=========================
A reusable component for running Excel analyses from OpenMDAO. It uses `xlwings` to communicate with Excel and works through the following procedure:
1. Start Excel.
2. Open a pre-defined workbook file.
3. *Optional:* Run pre-calculation macros.
4. Set the input variables in corresponding ranges/cells.
5. *Optional:* Run main macros.
6. Read values of the output ranges/cells and assign to output variables.
7. *Optional:* Run post-calculation macros.
8. Close the workbook without saving it.

Features:
* Cells can be addressed by named ranges (use them!) or regular means (e.g. `C3`).
* Timeout handling by mercilessly killing the Excel process after a specified number of seconds.
* The Excel process will be kept alive between evaluations, unless the previous one timed out.
* The macro runner will try to catch and log macro errors, raising an OpenMDAO `AnalysisError`.
* Rudimentary logging using the Python logging system.

Non-features/pitfalls/known issues:
* "Trust access to the VBA project object model" must be enabled.
* MacOS has not been tested and will probably not work.
* Non-scalar values (i.e. ranges with more than one cell) have not been tested and will probably not work.
* Multi-processing (i.e. MPI) has not been tested, but might work.
* Excel has a few numeric quirks and you are encouraged to be very critical of your results. I personally don't even recommend using Excel in this context, unless you *really* have to.
* The timeout handler runs in a separate Python thread. This shouldn't be a problem, but you might want to know.

# Compatibility
The component is regularly-ish tested on the latest stable versions of:
* Windows 10, 64-bit
* Microsoft Excel, 32-bit
* Python 3
* OpenMDAO

I wish I could set up a CI pipeline to make this process more transparent and repeatable, but Excel being a proprietary product makes this more or less impossible.

# Installation
```sh
pip install git+https://github.com/ovidner/openmdao-bridge-excel.git#egg=openmdao_bridge_excel
```

# Usage example
```python
import openmdao.api as om
from openmdao_bridge_excel import ExcelComponent, ExcelVar

prob = om.Problem()
...
excel_comp = prob.add_subsystem("excel", ExcelComponent(
  file_path="absolute/or/relative/path/to/file.xlsm",
  pre_macros=["foo", "bar"],
  main_macros=["main", "foo"],
  post_macros=["cleanup"],
  inputs=[
    ExcelVar("in_1", "Sheet1!C2"),
    ExcelVar("in_2", "NamedRangesAreSupported"),
  ],
  outputs=[
    ExcelVar("out_1", "B1"),
    ExcelVar("out_2", "AnotherNamedRange"),
  ],
  timeout=60,  # One minute
))
...
```
Input/output variables can then be addressed on the OpenMDAO component as usual (e.g. `excel.in_1`, `excel.out_2`).

# Development environment setup
This is automagically handled with `anaconda-project`:
```sh
anaconda-project run setup
```

You should then be able to run the tests:
```sh
anaconda-project run pytest
```
