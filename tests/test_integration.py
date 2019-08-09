import hypothesis.extra.numpy as np_st
import hypothesis.strategies as st
import numpy as np
import openmdao.api as om
import pytest
from hypothesis import assume, given, settings

from openmdao_bridge_excel import ExcelComponent
from openmdao_utils.external_tools import VarMap


@settings(deadline=1000)
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

    prob.setup()
    prob.run_model()
    prob.cleanup()

    # Using a normal == comparison will not consider NaNs as equal.
    assert np.allclose(prob["indeps.x"], value, atol=0.0, rtol=0.0, equal_nan=True)
    assert np.allclose(
        prob["passthrough.out"], prob["indeps.x"], atol=0.0, rtol=0.0, equal_nan=True
    )
