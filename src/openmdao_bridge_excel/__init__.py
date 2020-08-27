import dataclasses
import itertools
import os.path

import numpy as np
import openmdao.api as om
import xlwings
from pywintypes import com_error

from .macro_execution import run_and_raise_macro, wrap_macros
from .timeout_utils import TimeoutComponentMixin, kill_pid


def nans(shape):
    return np.ones(shape) * np.nan


@dataclasses.dataclass(frozen=True)
class ExcelVar:
    name: str
    range: str
    shape = (1,)


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
        for var in self.options["inputs"]:
            self.add_input(name=var.name, val=nans(var.shape))

        for var in self.options["outputs"]:
            self.add_output(name=var.name, val=nans(var.shape))

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
        for var in self.options["inputs"]:
            self.app.range(var.range).options(convert=np.array).value = inputs[var.name]

        self.app.calculation = "automatic"
        self.app.calculate()

        for macro in self.options["main_macros"]:
            run_and_raise_macro(book, macro, "main")

        for var in self.options["outputs"]:
            outputs[var.name] = (
                self.app.range(var.range).options(convert=np.array).value
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
