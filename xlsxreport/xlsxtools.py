def set_col_format(**args):
    """
    Encapsulate the column format dictionary, Options:
        width=Column with (xls with units)
        type=Python cell's data type
        title=Col's title string
        format=Cell format see: set_cell_format
        formula=Excel formula
    """
    return args


def set_cell_format(**args):
    """
    Encapsulate the cell format dictionary, Options:
        num_format='@',
        font_size=10,
        font_color='gray',
        bg_color='#FFEB9C'
        align='center',
        valign='vcenter',
        bold=True,
        text_wrap=True,
    """
    return args