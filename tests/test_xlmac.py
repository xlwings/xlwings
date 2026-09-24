import sys
from unittest.mock import MagicMock, call

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "darwin", reason="macOS backend only")


@pytest.mark.parametrize(
    ("pid_info", "expected"),
    [
        ('"pid"=2584\n', 2584),
        (
            "[ NULL ]  [ NULL ]  \n"
            "    bundleID=[ NULL ] \n"
            "    bundle path=[ NULL ] \n"
            "    executable path=[ NULL ] \n"
            "    pid = 15883 !cgsConnection !signalled type=[ NULL ]  "
            "flavor=[ NULL ]  Version=[ NULL ]  Arch=!!none \n",
            15883,
        ),
        ('"pid"=[ NULL ] \n', None),
        ("    pid = 0x1092 \n", None),
        ("    pid = 1503 token=[sess=100021 pid=1503]\n", 1503),
        ("    token=[sess=100021 pid=999]\n    pid = 1503\n", 1503),
        ("    bundleID=[ NULL ] \n", None),
    ],
)
def test_parse_pid(pid_info, expected):
    from xlwings._xlmac import _parse_pid

    assert _parse_pid(pid_info) == expected


def test_iter_excel_instances_accepts_macos_27_output(monkeypatch):
    from xlwings import _xlmac

    outputs = iter(
        [
            b"ASN:0x0-0x12345-Microsoft_Excel ",
            (
                b"[ NULL ]  [ NULL ]  \n"
                b"    bundleID=[ NULL ] \n"
                b"    bundle path=[ NULL ] \n"
                b"    executable path=[ NULL ] \n"
                b"    pid = 15883 !cgsConnection !signalled type=[ NULL ]\n"
            ),
        ]
    )
    monkeypatch.setattr(
        _xlmac.subprocess, "check_output", lambda command: next(outputs)
    )

    assert list(_xlmac.Apps()._iter_excel_instances()) == [15883]


def test_activate_accepts_macos_27_output(monkeypatch):
    from xlwings import _xlmac

    outputs = iter(
        [
            b"ASN:0x0-0x12345-com_example_Frontmost ",
            b"[ NULL ]  [ NULL ]  \n    bundleID=[ NULL ] \n    pid = 1234\n",
        ]
    )
    system_events = MagicMock()
    monkeypatch.setattr(
        _xlmac.subprocess, "check_output", lambda command: next(outputs)
    )
    monkeypatch.setattr(_xlmac.appscript, "app", lambda name: system_events)
    monkeypatch.setattr(_xlmac.App, "pid", property(lambda self: 5678))
    app = _xlmac.App.__new__(_xlmac.App)

    app.activate()

    assert system_events.processes.__getitem__.call_count == 2
    assert (
        system_events.processes.__getitem__.return_value.frontmost.set.call_args_list
        == [
            call(True),
            call(True),
        ]
    )
