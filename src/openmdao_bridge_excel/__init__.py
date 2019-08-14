import os.path

import numpy as np
import openmdao.api as om
import xlwings


def nans(shape):
    return np.ones(shape) * np.nan


class ExcelComponent(om.ExplicitComponent):
    def initialize(self):
        self.options.declare("file_path", types=str)
        self.options.declare("inputs", types=list)
        self.options.declare("outputs", types=list)
        self.options.declare("pre_macros", types=list, default=[])
        self.options.declare("main_macros", types=list, default=[])
        self.options.declare("post_macros", types=list, default=[])

        self.app = None

    def setup(self):
        for var_map in self.options["inputs"]:
            self.add_input(name=var_map.name, val=nans(var_map.shape))

        for var_map in self.options["outputs"]:
            self.add_output(name=var_map.name, val=nans(var_map.shape))

        self.app = xlwings.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False

    def compute(self, inputs, outputs, discrete_inputs=None, discrete_outputs=None):
        book = self.app.books.open(self.options["file_path"])

        for macro in self.options["pre_macros"]:
            book.macro(macro).run()

        self.app.calculation = "manual"
        for var_map in self.options["inputs"]:
            self.app.range(var_map.ext_name).options(convert=np.array).value = inputs[
                var_map.name
            ]

        self.app.calculation = "automatic"
        self.app.calculate()

        for macro in self.options["main_macros"]:
            book.macro(macro).run()

        for var_map in self.options["outputs"]:
            outputs[var_map.name] = (
                self.app.range(var_map.ext_name).options(convert=np.array).value
            )

        for macro in self.options["post_macros"]:
            book.macro(macro).run()

        # Closes without saving
        book.close()

    def cleanup(self):
        self.app.quit()
        super().cleanup()
