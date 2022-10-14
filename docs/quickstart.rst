Quickstart
==========

Please follow these steps to generate your first XlsxReport
Check :doc:`installation` to install XlsxReport first.


A Minimal Report
----------------

A minimal XlsxReport looks something like this:

.. code-block:: python

    from xlsxreport import XlsxReport, XlsxPage, XlsxTable

    # Create report object (Excel's file)
    rpt = XlsxReport('excel_file_name.xlsx')

    # Create & assign Workbook page
    page = XlsxPage('Sheet1')
    rpt.add_book_page('page', page)

    # Setup table data (Any list of dict)
    data = [
        {'First Name': 'John', 'Last Name': 'Smith', 'Age': 39},
        {'First Name': 'Mary', 'Last Name': 'Jane', 'Age': 25},
        {'First Name': 'Jennifer', 'Last Name': 'Doe', 'Age': 28},
    ]

    # Create & assign WorkSheet's table
    table = XlsxTable(data)
    page.add_page_table('table', table)

    # Define table layout (Use same cols as in data)
    table.cols_setup = {
        'First Name': {
            'type': str,
            'width': 15,
        },
        'Last Name': {
            'type': str,
            'width': 15,
        },
        'Age': {
            'type': int,
            'width': 10,
        },
    }

    # Create Excel's file
    rpt.generate()

