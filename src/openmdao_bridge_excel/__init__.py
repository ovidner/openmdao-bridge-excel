import itertools
import os.path

import numpy as np
import openmdao.api as om
import xlwings
from pywintypes import com_error
from openmdao_utils.external_tools import kill_pid
from openmdao_utils.timeout import TimeoutComponentMixin

from .macro_execution import run_and_raise_macro, wrap_macros


def nans(shape):
    return np.ones(shape) * np.nan


class ExcelComponent(TimeoutComponentMixin, om.ExplicitComponent):
    def initialize(self):
        self.options.declare("file_path", types=str)
        self.options.declare("inputs", types=list)
        self.options.declare("outputs", types=list)
        self.options.declare("pre_macros", types=list, default=[])
        self.options.declare("main_macros", types=list, default=[])
        self.options.declare("post_macros", types=list, default=[])

        self.app = None
        self.app_pid = None

    def setup(self):
        for var_map in self.options["inputs"]:
            self.add_input(name=var_map.name, val=nans(var_map.shape))

        for var_map in self.options["outputs"]:
            self.add_output(name=var_map.name, val=nans(var_map.shape))

        self.app = xlwings.App(visible=False, add_book=False)
        self.app_pid = self.app.pid
        self.app.display_alerts = False
        self.app.screen_updating = False

    def open_and_run(self, inputs, outputs, discrete_inputs, discrete_outputs):
        book = self.app.books.open(self.options["file_path"])

        all_macros = set(
            itertools.chain(
                self.options["pre_macros"],
                self.options["main_macros"],
                self.options["post_macros"],
            )
        )

        if len(all_macros):
            wrap_macros(book, all_macros)

        for macro in self.options["pre_macros"]:
            run_and_raise_macro(book, macro, "pre")

        self.app.calculation = "manual"
        for var_map in self.options["inputs"]:
            self.app.range(var_map.ext_name).options(convert=np.array).value = inputs[
                var_map.name
            ]

        self.app.calculation = "automatic"
        self.app.calculate()

        for macro in self.options["main_macros"]:
            run_and_raise_macro(book, macro, "main")

        for var_map in self.options["outputs"]:
            outputs[var_map.name] = (
                self.app.range(var_map.ext_name).options(convert=np.array).value
            )

        for macro in self.options["post_macros"]:
            run_and_raise_macro(book, macro, "post")

        # Closes without saving
        book.close()

    def compute(self, inputs, outputs, discrete_inputs=None, discrete_outputs=None):
        try:
            self.open_and_run(
                inputs, outputs, discrete_inputs or {}, discrete_outputs or {}
            )
        except Exception as exc:
            if self.timeout_state.reached:
                raise om.AnalysisError("Timeout reached!")
            else:
                raise exc

    def handle_timeout(self):
        kill_pid(self.app_pid)

    def cleanup(self):
        try:
            self.app.quit()
        except com_error as exc:
            pass
        kill_pid(self.app_pid)
        super().cleanup()
