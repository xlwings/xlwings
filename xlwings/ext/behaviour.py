# coding=utf-8
#
# Author: Sébastien de Menten <sdementen@gmail.com>
#

import contextlib

from ..constants import Calculation
from ..main import Application
from .monkey_patch import monkey_patch_register

@contextlib.contextmanager
def freeze(app, calculation=False, events=False, screen=False, alerts=False):
    """
    Context manager that freezes the Excel application regarding different behaviors.
    If a keyword is set to False (the default), the behavior is not changed. If it was already disabled, it will not be enabled.

    Arguments
    ---------
    app : Application
        The Excel application to freeze


    Keyword Arguments
    -----------------
    calculation: boolean, default False
        True if calculation must be set to manual within the context.

    events: boolean, default False
        True if Excel must not handle events within the context.

    screen: boolean, default False
        True if Excel must not update the screen within the context.

    alerts: boolean, default False
        True if Excel should not display alerts within the context.


    Example
    -------
    >>> from xlwings import Workbook, Application
    >>> wb = Workbook()
    >>> app = Application(wb)
    >>> # Excel will not recalculate or update the screen while processing the loop
    >>> with app.freeze(calculation=True, screen=True):
    >>>     for i in range(1000):
    >>>         Range((i,1)).value = i


    .. versionadded:: TODO: fill this
    """
    # save state of current behaviors
    save_state = (app.screen_updating,
                  app.enable_events,
                  app.calculation,
                  app.display_alerts,
                  )

    # if need to freeze the behavior, freeze it
    if events:
        app.enable_events = False
    if alerts:
        app.display_alerts = False
    if calculation:
        app.calculation = Calculation.xlCalculationManual
    if screen:
        app.screen_updating = False

    # yield in a try: finally: to force reset of behaviors even in case of exception
    try:
        yield
    finally:
        (app.screen_updating,
         app.enable_events,
         app.calculation,
         app.display_alerts,
         ) = save_state


# register the freeze function on the Application class (use enable_monkey_patch to enable it)
monkey_patch_register.append((Application, freeze))
