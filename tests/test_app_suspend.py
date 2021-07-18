import sys
from pathlib import Path

import pytest
import xlwings as xw

this_dir = Path(__file__).resolve().parent


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app


def test_suspend_empty(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
    with app.suspend():
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
                ('automatic', True, True, True))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))


def test_suspend_screen(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
    with app.suspend('screen'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating) ==
                ('automatic', True, True, False))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))


def test_suspend_events(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
    with app.suspend('events'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating) ==
                ('automatic', False, True, True))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))


def test_suspend_alerts(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
    with app.suspend('alerts'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating) ==
                ('automatic', True, False, True))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))


def test_suspend_calculation(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
    with app.suspend('calculation'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating) ==
                ('manual', True, True, True))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))


@pytest.mark.skipif(sys.platform.startswith('darwin'), reason='app.interactive is not supported on macOS')
def test_suspend_interactive(app):
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating, app.interactive),
            ('automatic', True, True, True, True))
    with app.suspend('interactive'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating, app.interactive) ==
                ('automatic', True, True, True, False))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True, True))


def test_suspend_multiple(app):
    with app.suspend('screen', 'events'):
        assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating) ==
                ('automatic', False, True, False))
    assert ((app.calculation, app.enable_events, app.display_alerts, app.screen_updating),
            ('automatic', True, True, True))
