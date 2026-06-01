"""
Tests for streaming_callback (xlwings Lite/Pyodide) path in custom_functions_call:
restart semantics, cancellation, single-emitter guarantee, and bookkeeping.
"""

import asyncio
import contextlib
from types import ModuleType

import pytest

from xlwings.pro.udfs_officejs import (
    background_tasks,
    custom_functions_call,
    task_key_to_sid_counts,
    task_key_to_task,
    xlfunc,
)


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.fixture(autouse=True)
async def _clean_state():
    """Reset module-level tracking dicts before each test."""
    background_tasks.clear()
    task_key_to_sid_counts.clear()
    task_key_to_task.clear()
    yield
    tasks = list(background_tasks.values())
    for task in tasks:
        task.cancel()
    await asyncio.sleep(0)
    background_tasks.clear()
    task_key_to_sid_counts.clear()
    task_key_to_task.clear()


def _make_module(func):
    """Wrap a decorated function into a fake module for custom_functions_call."""
    mod = ModuleType("test_module")
    setattr(mod, func.__name__, func)
    return mod


def _make_data(func_name, task_key, args=None):
    return {
        "func_name": func_name,
        "args": args or [],
        "task_key": task_key,
        "version": "placeholder",
        "client": "Office.js",
        "runtime": "1.4",
        "culture_info_name": "en-US",
        "date_format": "m/d/yy",
    }


class TestStreamingCallbackBasic:
    @pytest.mark.anyio
    async def test_streaming_callback_receives_results(self):
        results = []

        @xlfunc
        async def my_stream():
            yield 1
            yield 2
            yield 3

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(
            data, mod, streaming_callback=lambda r: results.append(r)
        )
        # Wait for the task to finish (finite generator)
        with contextlib.suppress(asyncio.CancelledError):
            await task

        assert len(results) == 3
        # Results come through convert() which wraps scalars in 2D lists
        assert results[0] == [[1]]
        assert results[1] == [[2]]
        assert results[2] == [[3]]

    @pytest.mark.anyio
    async def test_streaming_callback_creates_background_task(self):
        @xlfunc
        async def my_stream():
            while True:
                yield 1
                await asyncio.sleep(0.1)

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(data, mod, streaming_callback=lambda r: None)

        assert task is not None
        assert "my_stream_key" in background_tasks
        assert background_tasks["my_stream_key"] is task

    @pytest.mark.anyio
    async def test_streaming_callback_error_sent_to_callback(self):
        results = []

        @xlfunc
        async def my_stream():
            yield 1
            raise ValueError("test error")

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(
            data, mod, streaming_callback=lambda r: results.append(r)
        )
        with contextlib.suppress(asyncio.CancelledError):
            await task

        assert len(results) == 2
        assert results[0] == [[1]]
        assert isinstance(results[1], list)
        assert "ERROR:" in results[1][0][0]
        assert "test error" in results[1][0][0]


class TestStreamingCallbackRestart:
    @pytest.mark.anyio
    async def test_restart_cancels_old_task(self):
        """Re-invoking with same task_key cancels the old task."""
        results_old = []
        results_new = []

        @xlfunc
        async def my_stream():
            while True:
                yield "old"
                await asyncio.sleep(0.01)

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        old_task = await custom_functions_call(
            data, mod, streaming_callback=lambda r: results_old.append(r)
        )
        # Let it yield a few times
        await asyncio.sleep(0.05)

        assert len(results_old) > 0
        assert old_task in background_tasks.values()

        # Now re-invoke with same task_key but new callback
        new_task = await custom_functions_call(
            data, mod, streaming_callback=lambda r: results_new.append(r)
        )

        assert old_task.cancelled() or old_task.done()
        assert new_task is not old_task
        assert background_tasks["my_stream_key"] is new_task

    @pytest.mark.anyio
    async def test_restart_only_new_task_emits(self):
        """After restart, only the new task pushes results."""
        results = []

        call_count = 0

        @xlfunc
        async def my_stream():
            nonlocal call_count
            call_count += 1
            my_id = call_count
            while True:
                yield my_id
                await asyncio.sleep(0.01)

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        # First invocation
        await custom_functions_call(
            data, mod, streaming_callback=lambda r: results.append(r)
        )
        await asyncio.sleep(0.05)

        # Restart
        results.clear()
        await custom_functions_call(
            data, mod, streaming_callback=lambda r: results.append(r)
        )
        await asyncio.sleep(0.05)

        # All results after restart should be from the new task (id=2)
        assert all(r == [[2]] for r in results)

    @pytest.mark.anyio
    async def test_restart_bookkeeping_consistent(self):
        """After restart, background_tasks has exactly one entry for the key."""

        @xlfunc
        async def my_stream():
            while True:
                yield 1
                await asyncio.sleep(0.1)

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        await custom_functions_call(data, mod, streaming_callback=lambda r: None)
        await custom_functions_call(data, mod, streaming_callback=lambda r: None)
        await custom_functions_call(data, mod, streaming_callback=lambda r: None)

        assert len([k for k in background_tasks if k == "my_stream_key"]) == 1


class TestStreamingCallbackCancel:
    @pytest.mark.anyio
    async def test_cancel_via_task_cancel(self):
        """Direct task.cancel() stops the streaming task."""

        @xlfunc
        async def my_stream():
            while True:
                yield 1
                await asyncio.sleep(0.01)

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(data, mod, streaming_callback=lambda r: None)
        await asyncio.sleep(0.03)

        task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await task

        assert task.done()

    @pytest.mark.anyio
    async def test_on_task_done_cleans_up(self):
        """Finished tasks are removed from background_tasks by the done callback."""

        @xlfunc
        async def my_stream():
            yield 1  # finite — will complete

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(data, mod, streaming_callback=lambda r: None)
        with contextlib.suppress(asyncio.CancelledError):
            await task
        # Let the done callback fire
        await asyncio.sleep(0)

        assert "my_stream_key" not in background_tasks


class TestStreamingContext:
    @pytest.mark.anyio
    async def test_streaming_context_wraps_generator(self):
        """streaming_context is active while the generator runs."""
        context_active = []

        class TrackingContext:
            active = False

            def __enter__(self):
                self.active = True
                return self

            def __exit__(self, *exc):
                self.active = False

        ctx = TrackingContext()

        @xlfunc
        async def my_stream():
            context_active.append(ctx.active)
            yield 1
            context_active.append(ctx.active)
            yield 2

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        task = await custom_functions_call(
            data, mod, streaming_callback=lambda r: None, streaming_context=ctx
        )
        with contextlib.suppress(asyncio.CancelledError):
            await task

        assert all(context_active)


class TestStreamingCallbackNoSio:
    @pytest.mark.anyio
    async def test_sio_path_not_used_when_callback_provided(self):
        """When streaming_callback is set, sio is not called."""

        @xlfunc
        async def my_stream():
            yield 1

        mod = _make_module(my_stream)
        data = _make_data("my_stream", "my_stream_key")

        sio_called = False

        class FakeSio:
            async def emit(self, *args, **kwargs):
                nonlocal sio_called
                sio_called = True

        task = await custom_functions_call(
            data,
            mod,
            sio=FakeSio(),
            streaming_callback=lambda r: None,
        )
        with contextlib.suppress(asyncio.CancelledError):
            await task

        assert not sio_called
