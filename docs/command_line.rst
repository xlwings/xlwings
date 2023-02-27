.. _command_line:

Command Line Client (CLI)
=========================

xlwings comes with a command line client. On Windows, type the commands into a Command Prompt or Anaconda Prompt, on Mac, type them into a Terminal. To get an overview of all commands, simply type ``xlwings`` and hit Enter:

.. code-block:: text

    addin               Run "xlwings addin install" to install the Excel add-
                        in (will be copied to the user's XLSTART folder).
                        Instead of "install" you can also use "update",
                        "remove" or "status". Note that this command may take
                        a while. You can install your custom add-in by
                        providing the name or path via the --file/-f flag,
                        e.g. "xlwings add-in install -f custom.xlam or copy
                        all Excel files in a directory to the XLSTART folder
                        by providing the path via the --dir flag." To install
                        the add-in for every user globally, use the --glob/-g
                        flag and run this command from an Elevated Command
                        Prompt.
                        (New in 0.6.0, the --dir flag was added in 0.24.8 and the
                        --glob flag in 0.28.4)
    quickstart          Run "xlwings quickstart myproject" to create a folder
                        called "myproject" in the current directory with an
                        Excel file and a Python file, ready to be used. Use
                        the "--standalone" flag to embed all VBA code in the
                        Excel file and make it work without the xlwings add-
                        in. Use "--fastapi" for creating a project that uses a
                        remote Python interpreter. Use "--addin --ribbon" to
                        create a template for a custom ribbon addin. Leave
                        away the "--ribbon" if you don't want a ribbon tab.
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
                        xlwings Google Apps Script module. Run "xlwings copy
                        vba" to copy the standalone xlwings VBA module. Run
                        "xlwings copy vba --addin" to copy the xlwings VBA
                        module for custom add-ins.
                        (New in 0.26.0, 'vba' added in 0.28.7)
    auth                Microsoft Azure AD: "xlwings auth azuread", see
                        https://docs.xlwings.org/en/stable/server_authentication.html
                        (New in 0.28.6)
    vba                 This functionality allows you to easily write VBA code
                        in an external editor: run "xlwings vba edit" to
                        update the VBA modules of the active workbook from
                        their local exports everytime you hit save. If you run
                        this the first time, the modules will be exported from
                        Excel into your current working directory. To
                        overwrite the local version of the modules with those
                        from Excel, run "xlwings vba export". To overwrite the
                        VBA modules in Excel with their local versions, run
                        "xlwings vba import". The "--file/-f" flag allows you
                        to specify a file path instead of using the active
                        Workbook. Requires "Trust access to the VBA project
                        object model" enabled. NOTE: Whenever you change
                        something in the VBA editor (such as the layout of a
                        form or the properties of a module), you have to run
                        "xlwings vba export".
                        (New in 0.26.3, changed in 0.27.0)
