"""
Tests for streaming function lifecycle: refcount tracking, cancel-task,
disconnect cleanup, and reconnect resubscription semantics.

These tests exercise the module-level state and functions in
xlwings.pro.udfs_officejs directly, without needing a live Socket.IO server.
"""

import asyncio

import pytest

from xlwings.pro.udfs_officejs import (
    background_tasks,
    sio_cancel_task,
    sio_disconnect,
    task_key_to_sid_counts,
    task_key_to_task,
)


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.fixture(autouse=True)
async def _clean_state():
    """Reset module-level tracking dicts before each test."""
    task_key_to_sid_counts.clear()
    task_key_to_task.clear()
    background_tasks.clear()
    yield
    tasks = list(task_key_to_task.values()) + list(background_tasks.values())
    for task in tasks:
        task.cancel()
    await asyncio.sleep(0)  # let cancellations finalize
    task_key_to_sid_counts.clear()
    task_key_to_task.clear()
    background_tasks.clear()


def _make_dummy_task(name="xlwings-test"):
    async def forever():
        await asyncio.sleep(3600)

    return asyncio.create_task(forever(), name=name)


def _subscribe(task_key, sid, task=None):
    """Simulate a subscription: increment refcount and optionally register a task."""
    sid_counts = task_key_to_sid_counts.setdefault(task_key, {})
    sid_counts[sid] = sid_counts.get(sid, 0) + 1
    if task is not None:
        task_key_to_task[task_key] = task
        background_tasks[task_key] = task


# --- Refcount tracking ---


class TestRefcountCancel:
    @pytest.mark.anyio
    async def test_single_subscribe_and_cancel(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)

        await sio_cancel_task("sid1", "stream_A")
        await asyncio.sleep(0)  # let event loop finalize cancellation

        assert "stream_A" not in task_key_to_sid_counts
        assert "stream_A" not in task_key_to_task
        assert task.cancelled()

    @pytest.mark.anyio
    async def test_same_sid_two_subscriptions_cancel_one(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid1")

        await sio_cancel_task("sid1", "stream_A")

        assert task_key_to_sid_counts["stream_A"]["sid1"] == 1
        assert not task.cancelled()

    @pytest.mark.anyio
    async def test_same_sid_two_subscriptions_cancel_both(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid1")

        await sio_cancel_task("sid1", "stream_A")
        await sio_cancel_task("sid1", "stream_A")
        await asyncio.sleep(0)

        assert "stream_A" not in task_key_to_sid_counts
        assert task.cancelled()

    @pytest.mark.anyio
    async def test_two_sids_cancel_one_keeps_stream(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid2")

        await sio_cancel_task("sid1", "stream_A")

        assert task_key_to_sid_counts["stream_A"] == {"sid2": 1}
        assert not task.cancelled()

    @pytest.mark.anyio
    async def test_two_sids_cancel_both_tears_down(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid2")

        await sio_cancel_task("sid1", "stream_A")
        await sio_cancel_task("sid2", "stream_A")
        await asyncio.sleep(0)

        assert "stream_A" not in task_key_to_sid_counts
        assert task.cancelled()

    @pytest.mark.anyio
    async def test_cancel_nonexistent_task_key_is_noop(self):
        await sio_cancel_task("sid1", "nonexistent")
        assert task_key_to_sid_counts == {}

    @pytest.mark.anyio
    async def test_cancel_nonexistent_sid_is_noop(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)

        await sio_cancel_task("sid_unknown", "stream_A")

        assert task_key_to_sid_counts["stream_A"] == {"sid1": 1}
        assert not task.cancelled()


# --- Disconnect cleanup ---


class TestDisconnect:
    @pytest.mark.anyio
    async def test_disconnect_removes_sid_from_all_streams(self):
        task_a = _make_dummy_task("xlwings-A")
        task_b = _make_dummy_task("xlwings-B")
        _subscribe("stream_A", "sid1", task=task_a)
        _subscribe("stream_B", "sid1", task=task_b)

        await sio_disconnect("sid1")

        assert "stream_A" not in task_key_to_sid_counts
        assert "stream_B" not in task_key_to_sid_counts
        assert task_a.cancelled()
        assert task_b.cancelled()

    @pytest.mark.anyio
    async def test_disconnect_preserves_other_sids(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid2")

        await sio_disconnect("sid1")

        assert task_key_to_sid_counts["stream_A"] == {"sid2": 1}
        assert not task.cancelled()

    @pytest.mark.anyio
    async def test_disconnect_drops_all_refcounts_for_sid(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)
        _subscribe("stream_A", "sid1")  # refcount = 2

        await sio_disconnect("sid1")

        assert "stream_A" not in task_key_to_sid_counts
        assert task.cancelled()

    @pytest.mark.anyio
    async def test_disconnect_unknown_sid_is_noop(self):
        task = _make_dummy_task()
        _subscribe("stream_A", "sid1", task=task)

        await sio_disconnect("sid_unknown")

        assert task_key_to_sid_counts["stream_A"] == {"sid1": 1}
        assert not task.cancelled()

    @pytest.mark.anyio
    async def test_disconnect_after_task_already_cleaned_up(self):
        """If on_task_done already cleaned up task_key_to_task, disconnect
        still removes sid_counts without crashing."""
        task_key_to_sid_counts["stream_A"] = {"sid1": 1}
        # task_key_to_task intentionally has no entry

        await sio_disconnect("sid1")

        assert "stream_A" not in task_key_to_sid_counts
