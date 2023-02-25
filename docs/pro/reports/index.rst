xlwings Reports
===============

.. toctree::
    :maxdepth: 2
    :hidden:

    reports
    components_filters
    markdown

This feature requires xlwings PRO.

xlwings Reports is a solution for template-based Excel and PDF reporting, making the generation of pixel-perfect factsheets really simple. xlwings Reports allows business users without Python knowledge to create and maintain Excel templates without having to rely on a Python developer after the initial setup has been done: xlwings Reports separates the Python code (pre- and post-processing) from the Excel template (layout/formatting).

xlwings Reports supports all commonly required components:

* **Text**: Easily format your text via Markdown syntax.
* **Tables (dynamic)**: Write pandas DataFrames to Excel cells and Excel tables and format them dynamically based on the number of rows.
* **Charts**: Use your favorite charting engine: Excel charts, Matplotlib, or Plotly.
* **Images**: You can include both raster (e.g., png) or vector (e.g., svg) graphics, including dynamically generated ones, e.g., QR codes or plots.
* **Multi-column Layout**: Split your content up into e.g. a classic two column layout by using Frames.
* **Single Template**: Generate reports in various languages, for various funds etc. based on a single template.
* **PDF Report**: Generate PDF reports automatically and "print" the reports on PDFs in your corporate layout for pixel-perfect results including headers, footers, backgrounds and borderless graphics.
* **Easy Pre-processing**: Since everything is based on Python, you can connect with literally any data source and clean it with pandas or some other library.
* **Easy Post-processing**: Again, with Python you're just a few lines of code away from sending an email with the reports as attachment or uploading the reports to your web server, S3 bucket etc.
