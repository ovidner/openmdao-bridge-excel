import dataclasses
import threading
from contextlib import contextmanager

import psutil


@dataclasses.dataclass(eq=False)
class TimeoutState:
    timer = None
    reached = False


@contextmanager
def timeout(seconds, timeout_reached_fn):
    state = TimeoutState()

    def _timeout_reached_fn():
        state.reached = True
        timeout_reached_fn()

    timer = threading.Timer(seconds, _timeout_reached_fn)
    state.timer = timer
    timer.start()

    try:
        yield state
    finally:
        timer.cancel()


class TimeoutComponentMixin:
    def _declare_options(self):
        super()._declare_options()
        self.options.declare("timeout", types=(int, float), default=(60 * 60))

    def _apply_nonlinear(self):
        with timeout(self.options["timeout"], self._handle_timeout) as timeout_state:
            self.timeout_state = timeout_state
            super()._apply_nonlinear()
        self.timeout_state = None

    def _solve_nonlinear(self):
        with timeout(self.options["timeout"], self._handle_timeout) as timeout_state:
            self.timeout_state = timeout_state
            super()._solve_nonlinear()
        self.timeout_state = None

    def _handle_timeout(self):
        # TODO: logging
        self.handle_timeout()

    def handle_timeout(self):
        raise NotImplementedError()


def kill_pid(pid):
    try:
        proc = psutil.Process(pid)
        proc.kill()
    except psutil.NoSuchProcess:
        pass
