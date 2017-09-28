import xlrd
import xlwt


import regex


def get_sheet_by_name(book, name):
    """Get a sheet by name from xlwt.Workbook, a strangely missing method.
    Returns None if no sheet with the given name is present.
    """
    # Note, we have to use exceptions for flow control because the
    # xlwt API is broken and gives us no other choice.
    try:
        for idx in itertools.count():
            sheet = book.get_sheet(idx)
            if sheet.name == name:
                return sheet
    except IndexError:
        return None


def is_row_empty(row):
    for value in row:
        if value.value:
            return False
    return True

def find_table_start(table, title):
    """ Finds subtable of given title in a given table
    table -- xlrd table sheet 
    title -- regex for table title; regex allowed
    returns index of table start
    """
    i=0
    matches = []
    for value in table.col_values(0):
        match = regex.match(title,
                str(value))
        if match:
            matches.append((i, match))
        i += 1
    if not matches:
        print("Error: no matches", matches)

    title_index = sorted(matches,
                         key=lambda match: sum(match[1].fuzzy_counts))[0][0]

    i = 1
    # index of column names row
    current_row = table.row(title_index + i)
    while is_row_empty(current_row):
        
        i += 1
        current_row = table.row(title_index + i)
    
    return title_index + i


def find_table_end(table, start, end_text):
    """Finds subtable end row by text of the last rows' first column
    table -- xlrd sheet 
    start -- subtable's first row 
    end_text -- last row's first column's text; regex allowed

    Return table's last row index.
    """
    width = table.row_len(start)
    i = start
    
    while not regex.match(end_text, str(table.col_values(0)[i])):
        i += 1
    assert(is_row_empty(table.row(i+1)))
    return i


def get_table_start_end(raw_xls_path, title, end_text, answer_xls_path=None):
    """Get subtable start and end row indeces
    """

    table = xlrd.open_workbook(raw_xls_path).sheet_by_index(0)
    start_index = find_table_start(table, title)
    end_index = find_table_end(table, start_index, end_text)

    
    return table, start_index, end_index

# I.1.1 - нерухомість
def table_to_answer_table(table, start_index, end_index, answer_workbook_path, sheet_name):

    """Append answer table width subcolumns from table"""
    answer_workbook = xlwt.open(answer_workbook_path)
    answer = answer_workbook.get_sheet_by_name(sheet_name)
    print(answer)






