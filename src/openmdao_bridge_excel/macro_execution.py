import dataclasses
import hashlib
import logging

import openmdao.api as om

logger = logging.getLogger(__package__)

MACRO_WRAPPER_BASE = """Option Private Module
Option Explicit"""

MACRO_WRAPPER_INSTANCE = """Function {wrapped_macro_name}()
    On Error Resume Next
    {macro_name}
    {wrapped_macro_name} = Array(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext, Err.LastDllError)
End Function"""


@dataclasses.dataclass
class MacroError:
    number: int
    source: str
    description: str
    help_file: str
    help_context: str
    last_dll_error: int


@dataclasses.dataclass
class MacroResult:
    error: MacroError

    @property
    def success(self):
        return self.error.number == 0


def wrapper_macro_name(macro):
    macro_hash = hashlib.md5(macro.encode("utf-8")).hexdigest()
    return f"wrapped_{macro_hash}"


def wrap_macros(book, macros):
    vbe = book.app.api.VBE
    vb_project = vbe.ActiveVBProject

    wrapped_macros_comp = vb_project.VBComponents.Add(1)
    wrapped_macros_comp.Name = "ombe_wrapped_macros"

    wrapped_macros_code = wrapped_macros_comp.CodeModule
    wrapped_macros_code.AddFromString(MACRO_WRAPPER_BASE)

    for macro in macros:
        wrapped_macros_code.AddFromString(
            MACRO_WRAPPER_INSTANCE.format(
                macro_name=macro, wrapped_macro_name=wrapper_macro_name(macro)
            )
        )


def run_wrapped_macro(book, macro_name):
    error = book.macro(wrapper_macro_name(macro_name)).run()
    return MacroResult(error=MacroError(*error))


def run_and_raise_macro(book, macro, stage):
    logger.info(f"Running macro {macro} at {stage} stage...")
    result = run_wrapped_macro(book, macro)
    logger.info(
        f"Finished running macro {macro} at {stage} stage with result: {result}"
    )

    if not result.success:
        raise om.AnalysisError(
            f'Excel macro "{macro}" executed in "{stage}" stage failed: {result.error}'
        )
