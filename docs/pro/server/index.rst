xlwings Server (self-hosted)
============================

.. toctree::
    :maxdepth: 2
    :hidden:

    server
    officejs_addins
    officejs_custom_functions
    server_authentication

xlwings Server is a self-hosted and privacy-compliant solution that turns the Python dependency into a web app running on *your own* server (in the form of a serverless function, a fully managed container, etc.). Unlike Microsoft's *Python in Excel* solution, xlwings Server is not restricted to Office 365 but also works with the permanent versions of Office such as Office 2016 and Office 2021. It can be used from various clients:

* **VBA**: Desktop Excel (Windows and macOS)
* **Office Scripts**: Desktop Excel (Windows and macOS) and Excel on the web
* **Office.js Add-ins**: Desktop Excel (Windows and macOS), Excel on the web, and Excel on iPad (Note that Office.js add-ins don't work with the legacy `xls` format)
* **Google Apps Scripts**: Google Sheets


At the moment, xlwings Server doesn't cover yet 100% of the xlwings API. The following attributes are missing at the moment. If you need them, please reach out so we can prioritize their implementation:

.. code-block:: none

    xlwings.App

        - cut_copy_mode
        - quit()
        - display_alerts
        - startup_path
        - calculate()
        - status_bar
        - path
        - version
        - screen_updating
        - interactive
        - enable_events
        - calculation

    xlwings.Book

        - to_pdf()
        - save()

    xlwings.Characters

        - font
        - text

    xlwings.Chart

        - set_source_data()
        - to_pdf()
        - parent
        - delete()
        - top
        - width
        - height
        - name
        - to_png()
        - left
        - chart_type

    xlwings.Charts

        - add()

    xlwings.Font

        - size
        - italic
        - color
        - name
        - bold

    xlwings.Note

        - delete()
        - text

    xlwings.PageSetup

        - print_area

    xlwings.Picture

        - top
        - left
        - lock_aspect_ratio

    xlwings.Range

        - hyperlink
        - formula
        - font
        - width
        - formula2
        - characters
        - to_png()
        - columns
        - height
        - formula_array
        - paste()
        - rows
        - note
        - merge_cells
        - row_height
        - get_address()
        - merge()
        - to_pdf()
        - autofill()
        - top
        - wrap_text
        - merge_area
        - column_width
        - copy_picture()
        - table
        - unmerge()
        - current_region
        - left

    xlwings.Shape

        - parent
        - delete()
        - font
        - top
        - scale_height()
        - activate()
        - width
        - index
        - text
        - height
        - characters
        - name
        - type
        - scale_width()
        - left

    xlwings.Sheet

        - page_setup
        - used_range
        - shapes
        - charts
        - autofit()
        - copy()
        - to_html()
        - select()
        - visible

    xlwings.Table

        - display_name
        - show_table_style_last_column
        - show_table_style_column_stripes
        - insert_row_range
        - show_table_style_first_column
        - show_table_style_row_stripes