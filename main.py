from xlsxreport import XlsxReport, XlsxPage, XlsxTable

sample = [
    {'name': 'Jack', 'age': 26, 'gender': 'male'},
    {'name': 'John', 'age': 30, 'gender': 'male'},
    {'name': 'Mary', 'age': 24, 'gender': 'female'},
    {'name': 'Peter', 'age': 21, 'gender': 'male'},
]

# Create report object
rpt = XlsxReport('output/excel_file.xlsx')

# Create & assign Workbook page
page1 = XlsxPage('Sample')
page1.titles = [
    'Sample Table',
    'By XlsxReport',
    ]

rpt.add_book_page('page1', page1)

# Create & assign WorkSheet table
table1 = XlsxTable(sample)
table1.start_at_row = 4

page1.add_page_table('table1', table1)


rpt.generate()
