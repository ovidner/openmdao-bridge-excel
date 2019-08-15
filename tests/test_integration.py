import hypothesis.extra.numpy as np_st
import hypothesis.strategies as st
import numpy as np
import openmdao.api as om
import pytest
from hypothesis import given, settings

from openmdao_bridge_excel import ExcelComponent
from openmdao_utils.external_tools import VarMap


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
