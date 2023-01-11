.. rst-class:: hide-header

Welcome to XlsxReport
=====================

Welcome to XlsxReport's documentation.

This module wraps XlsxWriter to simplify the generation of
native excel files from list of dict.

See the following sections for more information:

- :doc:`installation`
- :doc:`quickstart`
- :doc:`cell_format`

User's Guide
------------

To generate an excel report just define:

- XlsxReport: Excel file object (Workbook)
- XlsxPage: Excel sheet objet (Worksheet)
- XlsxTable: Excel table data from list of dict

Import classes

.. code-block:: python

    from XlsxReport import XlsxReport, XlsxPage, XlsxTable

Declare a XlsxReport report/file.

.. code-block:: python

    rpt = XlsxReport('excel_file_name.xlsx')

Create page and add to excel book.

.. code-block:: python

    page = XlsxPage('Sheet1')  # "Sheet1" is page's name (in tab)
    rpt.add_book_page('page', page)  # Add "page" to report

