.. _command_line:

Command Line Client (CLI)
=========================

xlwings comes with a command line client. On Windows, type the commands into a Command Prompt or Anaconda Prompt, on Mac, type them into a Terminal. To get an overview of all commands, simply type ``xlwings`` and hit Enter:

.. code-block:: text

    addin               Run "xlwings addin install" to install the Excel add-
                        in (will be copied to the XLSTART folder). Instead of
                        "install" you can also use "update", "remove" or
                        "status". Note that this command may take a while. Use
                        the "--unprotected" flag to install the add-in without
                        password protection. You can install your custom add-
                        in by providing the name or path via the --file flag,
                        e.g. "xlwings add-in install --file custom.xlam or
                        copy all Excel files in a directory to the XLSTART
                        folder by providing the path via the --dir flag."
                        (New in 0.6.0, the unprotected flag was added in 0.20.4
                        and the --dir flag in 0.24.8)
    quickstart          Run "xlwings quickstart myproject" to create a folder
                        called "myproject" in the current directory with an
                        Excel file and a Python file, ready to be used. Use
                        the "--standalone" flag to embed all VBA code in the
                        Excel file and make it work without the xlwings add-
                        in.
    runpython           macOS only: run "xlwings runpython install" if you
                        want to enable the RunPython calls without installing
                        the add-in. This will create the following file:
                        ~/Library/Application
                        Scripts/com.microsoft.Excel/xlwings.applescript
                        (new in 0.7.0)
    restapi             Use "xlwings restapi run" to run the xlwings REST API
                        via Flask dev server. Accepts "--host" and "--port" as
                        optional arguments.
    license             xlwings PRO: Use "xlwings license update -k KEY" where
                        "KEY" is your personal (trial) license key. This will
                        update ~/.xlwings/xlwings.conf with the LICENSE_KEY
                        entry. If you have a paid license, you can run
                        "xlwings license deploy" to create a deploy key. This
                        is not available for trial keys.
    config              Run "xlwings config create" to create the user config
                        file (~/.xlwings/xlwings.conf) which is where the
                        settings from the Ribbon add-in are stored. It will
                        configure the Python interpreter that you are running
                        this command with. To reset your configuration, run
                        this with the "--force" flag which will overwrite your
                        current configuration.
                        (New in 0.19.5)
    code                Run "xlwings code embed" to embed all Python modules
                        of the workbook's dir in your active Excel file. Use
                        the "--file" flag to only import a single file by
                        providing its path. Requires xlwings PRO.
                        (Changed in 0.23.4)
    permission          "xlwings permission cwd" prints a JSON string that can
                        be used to permission the execution of all modules in
                        the current working directory via GET request.
                        "xlwings permission book" does the same for code that
                        is embedded in the active workbook.
                        (New in 0.23.4)
    release             Run "xlwings release" to configure your active
                        workbook to work with a one-click installer for easy
                        deployment. Requires xlwings PRO.
                        (New in 0.23.4)
    copy                Run "xlwings copy os" to copy the xlwings Office
                        Scripts module. Run "xlwings copy gs" to copy the
                        xlwings Google Apps Script module.
                        (New in 0.26.0)
    vba                 This functionality allows you to easily write VBA code
                        in an external editor: run "xlwings vba edit" to
                        update the VBA modules of the active workbook from
                        their local exports everytime you hit save. If you run
                        this the first time, the modules will be automatically
                        exported from Excel. To overwrite the local version of
                        the modules with the one from Excel, run "xlwings vba
                        export". The "--file" flag allows you to specify a
                        file path instead of using the active Workbook.
                        Requires "Trust access to the VBA project object
                        model" enabled. WARNING: Don't use this tool if you're
                        actively changing the UserForm Layout as it currently
                        syncs back the exported state, which would overwrite any
                        changes that you made in Excel.
                        (New in 0.27.0)