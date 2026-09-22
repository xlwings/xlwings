"""Headless variant of appscript's ASTranslate.app, patched to use xlwings' own
Excel terminology (xlwings/mac_dict.py) instead of appscript's default dictionary.

## What this is for

ASTranslate (https://github.com/appscript/appscript, folder ASTranslate/) is a
small macOS app that compiles a snippet of AppleScript, runs it against the
target application, and prints the equivalent appscript/py-appscript Python
call for every Apple event AppleScript sends along the way. It's the standard
way to figure out how to phrase an appscript call for some AppleScript command
you already know works.

Its "Untranslated event" error only means ASTranslate's own Excel dictionary
doesn't have a name for that event code -- it says nothing about whether
appscript or Excel can actually handle it. ASTranslate builds its Excel
connection with a plain `appscript.app("Microsoft Excel")`, i.e. no `terms=`,
so it always uses appscript's own (possibly incomplete) dump of Excel's
dictionary. xlwings ships its own hand-maintained dictionary in
xlwings/mac_dict.py (see that file's docstring for why), which is more
complete for the commands xlwings' _xlmac.py engine actually uses -- e.g.
`get_axis` (Excel's "get axis" command, used to reach a chart's ChartAxis) is
present in xlwings/mac_dict.py but missing from ASTranslate's default
dictionary, so vanilla ASTranslate reports it as untranslated even though it
works fine once the right dictionary is loaded.

This script reimplements ASTranslate's translate-and-render pipeline without
its Cocoa GUI (so it can run from a terminal / be scripted), and swaps in
xlwings/mac_dict.py as the terminology. Nothing here ships in the xlwings
package; it's a standalone dev tool for investigating xlwings' macOS engine
(xlwings/_xlmac.py) against Excel's real AppleScript dictionary.

## One-time setup

This script depends on ASTranslate's C extension and renderer modules, which
live in the separate appscript repository, not in xlwings:

    git clone https://github.com/hhas/appscript

Then compile the `_astranslate` C extension (it wraps the AppleScript OSA
component; no PyObjC/py2app/GUI needed for this headless use):

    cd appscript/ASTranslate/src
    clang -bundle -undefined dynamic_lookup \\
        -I"$(python3 -c 'import sysconfig; print(sysconfig.get_paths()["include"])')" \\
        -DMAC_OS_X_VERSION_MIN_REQUIRED=101200 \\
        -framework Carbon \\
        -o _astranslate"$(python3 -c 'import sysconfig; print(sysconfig.get_config_var("EXT_SUFFIX"))')" \\
        _astranslate.c

Set ASTRANSLATE_SRC below (or the ASTRANSLATE_SRC env var) to that
`appscript/ASTranslate/src` directory.

This needs the same Python environment xlwings' macOS engine runs under
(appscript, aem, etc. -- see xlwings' macOS extra), run with the system
Python or a venv that has appscript installed, not one that's missing Carbon
framework access.

## Usage

    python3 scripts/astranslate_headless.py path/to/snippet.applescript

Prints each translated Apple event as py-appscript code (mirroring
ASTranslate's Python tab), followed by the script's overall result.

Passing `send_events=False` to `translate()` mirrors ASTranslate's "Send
events to app" checkbox being unchecked: Apple events are still compiled and
sniffed for translation, but never actually sent, so a command's return value
is unavailable to later parts of the script (AppleScript will error on that
part) -- useful for previewing destructive commands. Left on by default here
since inspecting read/write chart-axis calls requires seeing real results.
"""

import importlib.util
import os
import sys
import traceback

ASTRANSLATE_SRC = os.environ.get(
    "ASTRANSLATE_SRC",
    os.path.expanduser("~/dev/appscript/ASTranslate/src"),
)
MAC_DICT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), os.pardir, "xlwings", "mac_dict.py"
)

sys.path.insert(0, ASTRANSLATE_SRC)

import _astranslate  # noqa: E402  (the compiled C extension, see setup above)
import aem  # noqa: E402
import appscript  # noqa: E402
import appscript.reference  # noqa: E402
import pythonrenderer  # noqa: E402  (from ASTranslate/src)
from aem import ae, kae  # noqa: E402
from constants import (  # noqa: E402
    UntranslatedKeywordError,
    UntranslatedUserPropertyError,
    kNoParam,
)

# Load xlwings' own Excel terminology instead of appscript's default dump, so
# commands like get_axis (event code sCRTXGAx) are recognized.
_spec = importlib.util.spec_from_file_location("mac_dict", MAC_DICT_PATH)
mac_dict = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mac_dict)

_standard_codecs = aem.Codecs()
# Keyed by (addressdesc.type, addressdesc.data): reuse one appscript.app() /
# terminology lookup per target application across all events in a script,
# same as ASTranslate's eventformatter._appCache.
_app_cache = {}


def _unpack_event_attributes(event):
    atts = []
    for code in [kae.keyEventClassAttr, kae.keyEventIDAttr, kae.keyAddressAttr]:
        atts.append(_standard_codecs.unpack(event.getattr(code, kae.typeWildCard)))
    return atts[0].code + atts[1].code, atts[2]


def _make_send_proc(is_live):
    """Build the OSASendProc AppleScript calls for every Apple event it sends.

    This is a trimmed, GUI-free copy of ASTranslate's
    eventformatter.makeCustomSendProc: same translation logic, but printing
    to stdout instead of a Cocoa text view, and importing only the Python
    renderer (the Ruby/Node renderers need py-osaterminology, which isn't
    needed for this).
    """

    def custom_send_proc(event, mode_flags, timeout):
        eventcode = None
        try:
            eventcode, addressdesc = _unpack_event_attributes(event)
            app_path = ae.addressdesctopath(addressdesc)

            if (addressdesc.type, addressdesc.data) not in _app_cache:
                if addressdesc.type != kae.typeProcessSerialNumber:
                    raise RuntimeError(
                        "Can't identify application (addressdesc descriptor "
                        "not typeProcessSerialNumber)"
                    )
                # The one change from ASTranslate's own eventformatter.py:
                # pass terms=mac_dict so Excel-specific commands are known.
                app = appscript.app(app_path, terms=mac_dict)
                app_data = app.AS_appdata
                _app_cache[(addressdesc.type, addressdesc.data)] = (app, app_data)
            app, app_data = _app_cache[(addressdesc.type, addressdesc.data)]

            desc = event.coerce(kae.typeAERecord)
            params = {}
            for i in range(desc.count()):
                key, value = desc.getitem(i + 1, kae.typeWildCard)
                params[key] = app_data.unpack(value)
            result_type = params.pop(b"rtyp", None)
            direct_param = params.pop(b"----", kNoParam)
            try:
                subject = app_data.unpack(
                    event.getattr(kae.keySubjectAttr, kae.typeWildCard)
                )
            except Exception:
                subject = None

            if subject is not None:
                target_ref = subject
            elif eventcode == b"coresetd":
                # 'set' command: direct param is the target, 'data' is the value
                target_ref = direct_param
                direct_param = params.pop(b"data")
            elif isinstance(direct_param, appscript.reference.Reference):
                target_ref = direct_param
                direct_param = kNoParam
            else:
                target_ref = app

            _timeout = timeout
            if timeout == 0xFFFFFFFF or timeout > 0x888888888888:
                _timeout = -1  # kAEDefaultTimeout, mis-sized as 64-bit
            elif timeout == 0xFFFFFFFE:
                _timeout = -2  # kNoTimeOut, mis-sized as 64-bit

            try:
                print(
                    pythonrenderer.renderCommand(
                        app_path,
                        addressdesc,
                        eventcode,
                        target_ref,
                        direct_param,
                        params,
                        result_type,
                        mode_flags,
                        _timeout,
                        app_data,
                    )
                )
            except (UntranslatedKeywordError, UntranslatedUserPropertyError) as e:
                print(f"Untranslated event {eventcode!r}\n{e}")
            except Exception:
                traceback.print_exc()
                print(f"Untranslated event {eventcode!r}")

        except Exception:
            traceback.print_exc()
            print(f"Untranslated event {eventcode!r}")

        if is_live:
            return event.send(mode_flags, timeout)
        else:
            return ae.newdesc(kae.typeNull, b"")

    return custom_send_proc


def translate(source, send_events=True):
    """Compile and run `source` (an AppleScript snippet), printing the
    py-appscript translation of every Apple event it sends, followed by the
    script's own result or error.
    """
    source_desc = _standard_codecs.pack(source)
    handler = _make_send_proc(is_live=send_events)
    result = _astranslate.translate(source_desc, handler)
    if result[0]:
        script = _standard_codecs.unpack(result[1])
        print("--- OK ---")
        print(script)
    else:
        script, error_num, error_msg, pos = (
            _standard_codecs.unpack(d) for d in result[1:]
        )
        kind = "Runtime" if script else "Compilation"
        print(f"--- {kind} Error ---")
        print(f"{error_msg} ({error_num})")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(f"Usage: python3 {sys.argv[0]} path/to/snippet.applescript")
    with open(sys.argv[1]) as f:
        translate(f.read())
