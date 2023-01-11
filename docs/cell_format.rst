Formating table cells
=====================

Cell's are formatted in **table.cols_setup** dict, you can define one or more attributes
as you need:

width
-----

Column with int (xls with units)

type
----

 Python cell's data type: str, float, int ...


 title
 -----

 Col's title

 format
 ------

 Set cell font attr:

 - num_format
 - font_size
 - font_color
 - bg_color
 - align
 - valign
 - bold
 - text_wrap': True,

formula
-------


internal_link
-------------

Can be used to create an internal link to other book's page when cell value = page name

Must indicate a destination cell: 'A2'

