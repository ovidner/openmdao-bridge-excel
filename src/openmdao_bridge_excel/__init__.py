import os.path

import numpy as np
import openmdao.api as om
import pythoncom
from win32com.client import gencache


def nans(shape):
    return np.ones(shape) * np.nan


def start_excel():
    clsid = pythoncom.CoCreateInstanceEx(
        "Excel.Application",
        None,
        pythoncom.CLSCTX_SERVER,
        None,
        (pythoncom.IID_IDispatch,),
    )[0]
    if gencache.is_readonly:
        # fix for "freezed" app: py2exe.org/index.cgi/UsingEnsureDispatch
        gencache.is_readonly = False
        gencache.Rebuild()

    return gencache.EnsureDispatch(clsid)


class ExcelComponent(om.ExplicitComponent):
    def initialize(self):
        self.options.declare("file_path", types=str)
        self.options.declare("inputs", types=list)
        self.options.declare("outputs", types=list)

        self.application = None

    def setup(self):
        for var_map in self.options["inputs"]:
            self.add_input(name=var_map.name, val=nans(var_map.shape))

        for var_map in self.options["outputs"]:
            self.add_output(name=var_map.name, val=nans(var_map.shape))

        self.application = start_excel()
        self.application.Visible = False
        self.application.Interactive = False
        self.application.DisplayAlerts = False
        self.application.EnableSound = False
        self.application.ScreenUpdating = False

    def compute(self, inputs, outputs, discrete_inputs=None, discrete_outputs=None):
        file_path = os.path.abspath(self.options["file_path"])
        workbook = self.application.Workbooks.Open(file_path, 3, True)

        # Disables automatic calculation
        self.application.Calculation = -4135
        for var_map in self.options["inputs"]:
            self.application.Range(var_map.ext_name).Value = inputs[
                var_map.name
            ].tolist()

        self.application.Calculate()

        # TODO: Here is a good time to run some macros.

        for var_map in self.options["outputs"]:
            range_ = self.application.Range(var_map.ext_name)
            if self.application.WorksheetFunction.IsError(range_):
                raise om.AnalysisError()
            outputs[var_map.name] = np.array(range_.Value)

        # Closes without saving
        workbook.Close(False)

    def cleanup(self):
        self.application.Quit()
        self.application = None
        super().cleanup()
