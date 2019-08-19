import dataclasses
import time

import hypothesis.extra.numpy as np_st
import hypothesis.strategies as st
import numpy as np
import openmdao.api as om
import pytest
from hypothesis import given, settings

from openmdao_bridge_excel import ExcelComponent
from openmdao_utils.external_tools import VarMap


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


@given(st.floats(allow_nan=False, allow_infinity=False))
def test_continuous_finite_scalar(value):
    prob = om.Problem()
    model = prob.model

    model.add_subsystem("indeps", om.IndepVarComp("x", val=value))
    model.add_subsystem(
        "passthrough",
        ExcelComponent(
            file_path="tests/data/passthrough.xlsx",
            inputs=[VarMap("in", "a")],
            outputs=[VarMap("out", "b")],
        ),
    )
    model.connect("indeps.x", "passthrough.in")

    try:
        prob.setup()
        prob.run_model()
    finally:
        prob.cleanup()

    # Using a normal == comparison will not consider NaNs as equal.
    assert np.allclose(prob["indeps.x"], value, atol=0.0, rtol=0.0, equal_nan=True)
    assert np.allclose(
        prob["passthrough.out"], prob["indeps.x"], atol=0.0, rtol=0.0, equal_nan=True
    )


@given(st.floats(allow_nan=False, allow_infinity=False))
def test_continuous_finite_scalar_macros(value):
    prob = om.Problem()
    model = prob.model

    model.add_subsystem("indeps", om.IndepVarComp("x", val=value))
    model.add_subsystem(
        "passthrough",
        ExcelComponent(
            file_path="tests/data/passthrough_macros.xlsm",
            inputs=[VarMap("in", "a")],
            outputs=[VarMap("out", "b")],
            pre_macros=["NameA", "NameB"],
            main_macros=["CopyAToB"],
            post_macros=["EnsureBEqualsA"],
        ),
    )
    model.connect("indeps.x", "passthrough.in")

    try:
        prob.setup()
        prob.run_model()
    finally:
        prob.cleanup()

    # Using a normal == comparison will not consider NaNs as equal.
    assert np.allclose(prob["indeps.x"], value, atol=0.0, rtol=0.0, equal_nan=True)
    assert np.allclose(
        prob["passthrough.out"], prob["indeps.x"], atol=0.0, rtol=0.0, equal_nan=True
    )


@pytest.mark.parametrize("stage", ["pre", "main", "post"])
def test_macro_errors(stage):
    fudge_up_macros = ["FudgeUp"]
    prob = om.Problem()
    model = prob.model

    model.add_subsystem("indeps", om.IndepVarComp("x", val=1))
    model.add_subsystem(
        "passthrough",
        ExcelComponent(
            file_path="tests/data/fudge_up.xlsm",
            inputs=[VarMap("in", "A1")],
            outputs=[VarMap("out", "A1")],
            pre_macros=fudge_up_macros if stage == "pre" else [],
            main_macros=fudge_up_macros if stage == "main" else [],
            post_macros=fudge_up_macros if stage == "post" else [],
        ),
    )
    model.connect("indeps.x", "passthrough.in")

    try:
        prob.setup()
        with pytest.raises(
            om.AnalysisError,
            match=f'Excel macro "FudgeUp" executed in "{stage}" stage failed',
        ):
            prob.run_model()
    finally:
        prob.cleanup()


@given(
    timeout=st.floats(min_value=2, max_value=8, allow_nan=False, allow_infinity=False)
)
@settings(deadline=10000, max_examples=3)
@pytest.mark.parametrize("stage", ["pre", "main", "post"])
@pytest.mark.parametrize("slow_macros", [["Breakable10"], ["NonBreakable10"]])
def test_timeout(stage, timeout, slow_macros):
    prob = om.Problem()
    model = prob.model

    model.add_subsystem("indeps", om.IndepVarComp("x", val=1))
    model.add_subsystem(
        "passthrough",
        ExcelComponent(
            file_path="tests/data/sleep.xlsm",
            inputs=[VarMap("in", "A1")],
            outputs=[VarMap("out", "A1")],
            pre_macros=slow_macros if stage == "pre" else [],
            main_macros=slow_macros if stage == "main" else [],
            post_macros=slow_macros if stage == "post" else [],
            timeout=timeout,
        ),
    )
    model.connect("indeps.x", "passthrough.in")

    try:
        prob.setup()
        with pytest.raises(om.AnalysisError, match="Timeout reached!"):
            with ExecutionTime() as execution_time:
                prob.run_model()
    finally:
        prob.cleanup()

    # Should be finished within the timeout limit plus some overhead, but not too early
    assert timeout <= execution_time.duration <= (timeout + 1)
