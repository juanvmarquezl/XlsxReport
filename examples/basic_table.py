from XlsxReport import XlsxReport, XlsxPage, XlsxTable
import csv
import os


# Create report object
rpt = XlsxReport('basic_table.xlsx')

# Create & assign Workbook page
page1 = XlsxPage('Cars')
page1.titles = [
    'Basic Table Sample',
    'Cars',
    'A dataset of about 400 cars with 8 characteristics such as horsepower, acceleration, etc.',
    'https://perso.telecom-paristech.fr/eagan/class/igr204/datasets',
    'By XlsxReport',
    ]
rpt.add_book_page('page1', page1)

# load & assign WorkSheet table
file_name = os.path.join('examples', 'data', 'cars.csv')
with open(file_name, encoding='utf-8') as csv_file:
    csv_data = csv.DictReader(csv_file, delimiter=';')
    sample = [dict(x) for x in csv_data]

table1 = XlsxTable(sample)
table1.start_at_row = page1.after_page_titles()
table1.start_at_col = 2
table1.cols_setup = {
    'Car': {
        'type': str,
        'width': 30,
    },
    'MPG': {
        'type': float,
        'format': {'num_format': '#,##0.00'},
    },
    'Cylinders': {
        'type': int,
        'format': {'num_format': '#,##0',
                   'bold': True,
                   'align': 'center'},
    },
    'Displacement': {
        'type': float,
        'format': {'num_format': '#,##0.00'},
    },
    'Horsepower': {
        'type': float,
        'format': {'num_format': '#,##0.00'},
    },
    'Weight': {
        'type': float,
        'format': {'num_format': '#,##0.00'},
    },
    'HP/W Ratio': {
        'type': float,
        'format': {'num_format': '#,##0.00'},
        'formula': '{{Horsepower}}/{{Weight}}',
    },

    'Acceleration': {
        'type': float,
        'width': 11.5,
        'format': {'num_format': '#,##0.00'},
    },
    'Model': {
        'type': int,
        'width': 7,
        'format': {'num_format': '#,##0', 'align': 'center'},
    },
    'Origin': {
        'type': str,
        'width': 10,
        'format': {'align': 'center'},
    },
}

page1.add_page_table('table1', table1)


rpt.generate()