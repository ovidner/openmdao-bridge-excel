import dataclasses
import time

import hypothesis.extra.numpy as np_st
import hypothesis.strategies as st
import numpy as np
import openmdao.api as om
import pytest
from hypothesis import given, settings

from openmdao_bridge_excel import ExcelComponent, ExcelVar

TEST_FILE_PATH = "tests/data/test.xlsm"


@dataclasses.dataclass
class ExecutionTime:
    start: float = dataclasses.field(init=False)
    end: float = dataclasses.field(init=False)

    def __enter__(self):
        self.start = time.time()
        return self

    def __exit__(self, *args):
        self.end = time.time()

    @property
    def duration(self):
        return (self.end - self.start) if (self.start and self.end) else None


@settings(deadline=5000)
@given(st.floats(allow_nan=False, allow_infinity=False))
@pytest.mark.parametrize("mode", ["formula", "macro"])
def test_continuous_finite_scalar(mode, value):
    prob = om.Problem()
    model = prob.model

    if mode == "formula":
        comp = ExcelComponent(
            file_path=TEST_FILE_PATH,
            inputs=[ExcelVar("in", "FormulaA")],
            outputs=[ExcelVar("out", "FormulaB")],
        )
    elif mode == "macro":
        comp = ExcelComponent(
            file_path=TEST_FILE_PATH,
            inputs=[ExcelVar("in", "MacroA")],
            outputs=[ExcelVar("out", "MacroB")],
            pre_macros=["NameA", "NameB"],
            main_macros=["CopyAToB"],
            post_macros=["EnsureBEqualsA"],
        )
    else:
        raise ValueError(mode)

    model.add_subsystem(
        "excel", comp,
    )

    try:
        prob.setup()
        prob.set_val("excel.in", value)
        prob.run_model()
    finally:
        prob.cleanup()

    # Using a normal == comparison will not consider NaNs as equal.
    assert np.allclose(prob["excel.out"], value, atol=0.0, rtol=0.0, equal_nan=True)


@pytest.mark.parametrize("stage", ["pre", "main", "post"])
def test_macro_errors(stage):
    fudge_up_macros = ["FudgeUp"]
    prob = om.Problem()
    model = prob.model

    model.add_subsystem(
        "excel",
        ExcelComponent(
            file_path=TEST_FILE_PATH,
            inputs=[],
            outputs=[],
            pre_macros=fudge_up_macros if stage == "pre" else [],
            main_macros=fudge_up_macros if stage == "main" else [],
            post_macros=fudge_up_macros if stage == "post" else [],
        ),
    )

    try:
        prob.setup()
        with pytest.raises(
            om.AnalysisError,
            match=f'Excel macro "FudgeUp" executed in "{stage}" stage failed',
        ):
            prob.run_model()
    finally:
        prob.cleanup()


@pytest.mark.parametrize("stage", ["pre", "main", "post"])
@pytest.mark.parametrize("timeout", [1, 10])
@pytest.mark.parametrize("slow_macros", [["SleepBreakable"], ["SleepNonbreakable"]])
def test_timeout(stage, timeout, slow_macros):
    prob = om.Problem()
    model = prob.model

    model.add_subsystem(
        "excel",
        ExcelComponent(
            file_path=TEST_FILE_PATH,
            inputs=[],
            outputs=[ExcelVar("out", "A1")],
            pre_macros=slow_macros if stage == "pre" else [],
            main_macros=slow_macros if stage == "main" else [],
            post_macros=slow_macros if stage == "post" else [],
            timeout=timeout,
        ),
    )

    try:
        prob.setup()
        with pytest.raises(om.AnalysisError, match="Timeout reached!"):
            with ExecutionTime() as execution_time:
                prob.run_model()
    finally:
        prob.cleanup()

    # Should be finished within the timeout limit plus some overhead, but not too early
    assert timeout <= execution_time.duration <= (timeout + 3)


@pytest.mark.parametrize("stage", ["main", "post"])
@pytest.mark.parametrize("slow_macros", [["SleepBreakable"], ["SleepNonbreakable"]])
@pytest.mark.parametrize("value", [1, 3])
def test_timeout_recovery(stage, slow_macros, value):
    prob = om.Problem()
    model = prob.model

    comp = model.add_subsystem(
        "excel",
        ExcelComponent(
            file_path=TEST_FILE_PATH,
            inputs=[
                ExcelVar("in", "FormulaA"),
                ExcelVar("sleep_duration", "SleepDuration"),
            ],
            outputs=[ExcelVar("out", "FormulaB")],
            # We can't adjust the sleep duration of the pre stage, so we let it be.
            pre_macros=[],
            main_macros=slow_macros if stage == "main" else [],
            post_macros=slow_macros if stage == "post" else [],
            timeout=5,
        ),
    )

    try:
        prob.setup()

        prob.set_val("excel.in", value)
        prob.set_val("excel.sleep_duration", 60)
        with pytest.raises(om.AnalysisError, match="Timeout reached!"):
            prob.run_model()

        prob.set_val("excel.in", value)
        prob.set_val("excel.sleep_duration", 0)
        prob.run_model()
        assert prob.get_val("excel.out") == value
    finally:
        prob.cleanup()
