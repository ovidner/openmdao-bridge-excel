import hashlib
import itertools
import os.path

import numpy as np
import openmdao.api as om
import xlwings


def nans(shape):
    return np.ones(shape) * np.nan


MACRO_WRAPPER_BASE = """Option Private Module
Option Explicit"""

MACRO_WRAPPER_INSTANCE = """Function {wrapped_macro_name}()
    On Error Resume Next
    {macro_name}
    {wrapped_macro_name} = Array(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext, Err.LastDllError)
End Function"""


def wrapper_macro_name(macro_name):
    macro_name_hash = hashlib.md5(macro_name.encode("utf-8")).hexdigest()
    return f"wrapped_{macro_name_hash}"


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

        all_macros = set(
            itertools.chain(
                self.options["pre_macros"],
                self.options["main_macros"],
                self.options["post_macros"],
            )
        )

        if len(all_macros):
            vbe = self.app.api.VBE
            vb_project = vbe.ActiveVBProject

            wrapped_macros_comp = vb_project.VBComponents.Add(1)
            wrapped_macros_comp.Name = "ombe_wrapped_macros"

            wrapped_macros_code = wrapped_macros_comp.CodeModule
            wrapped_macros_code.AddFromString(MACRO_WRAPPER_BASE)

            for macro_name in all_macros:
                wrapped_macros_code.AddFromString(
                    MACRO_WRAPPER_INSTANCE.format(
                        macro_name=macro_name,
                        wrapped_macro_name=wrapper_macro_name(macro_name),
                    )
                )

        for macro in self.options["pre_macros"]:
            book.macro(wrapper_macro_name(macro)).run()

        self.app.calculation = "manual"
        for var_map in self.options["inputs"]:
            self.app.range(var_map.ext_name).options(convert=np.array).value = inputs[
                var_map.name
            ]

        self.app.calculation = "automatic"
        self.app.calculate()

        for macro in self.options["main_macros"]:
            book.macro(wrapper_macro_name(macro)).run()

        for var_map in self.options["outputs"]:
            outputs[var_map.name] = (
                self.app.range(var_map.ext_name).options(convert=np.array).value
            )

        for macro in self.options["post_macros"]:
            book.macro(wrapper_macro_name(macro)).run()

        # Closes without saving
        book.close()

    def cleanup(self):
        self.app.quit()
        super().cleanup()
