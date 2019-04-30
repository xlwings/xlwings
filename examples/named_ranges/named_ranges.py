import xlwings as xw

 
_book = 'named_ranges.xlsm'

@xw.sub()
def kill_broken_names():
    """ Deletes all of the named ranges in a book with broken refernces. """
    
    wb = vba_true_caller(xw.Book.caller()) # just in case this function is imported to xlam
    names = [name.name for name in wb.names]
    for name in names:
        if '#REF!' in wb.names[name].refers_to:
            del wb.names[name]

@xw.sub()
def name_used_range_sheet_active():
    """ Creates a named range (ur_ + sheet name) for the used range of the active sheet """
    wb = vba_true_caller(xw.Book.caller()) # just in case this function is imported to xlam
    name_used_range(wb,wb.sheets.active.name)

@xw.sub()
def name_used_range_sheets_all():
    """ Creates a named range (ur_ + sheet name) for the used range of all sheets """
    wb = vba_true_caller(xw.Book.caller()) # just in case this function is imported to xlam
    name_used_range(wb)


    
    
    
    
def vba_true_caller(caller):
    """ If the calling book is an addin, find the active book of the instance. """
    if caller.name[-5:] == '.xlam':
        return caller.app.books[caller.app.api.activeworkbook.name]
    else:
        return caller
    
    

def name_used_range(wb:xw.main.Book,sheet_name:str=None,first_call:bool=True):
    """ 
    When sheet_name is specified, this function names the used range of the 
    sheet to be "ur_" + the name of the sheet.
    
    When the sheet_name is not specified, this function is applied to all 
    sheets of the book.
    
    No name is applied to blank sheets.
    """
    if sheet_name is None:
        for sheet in wb.sheets:
            name_used_range(wb,sheet_name=sheet.name)
    else:
        rng = used_range(wb.sheets[sheet_name])
        name = 'ur_'+sheet_name
        
        if rng is None: 
            if name in wb.names:
                del wb.names[name]        
        else:
            if name in wb.names:
                wb.names[name].refers_to = rng['r1c1']['full']
            else:
                wb.names.add(name,rng['r1c1']['full'])
            

    


def used_range(sht:xw.main.Sheet):
    
    """
        This function finds the used range of a sheet.  It returns the 
        range in three ways:
        
        It returns the range in a1 format with and without the sheet name.
        It returns the range in r1c1 format with and without the sheet name.
        It returns a row,column pair for the bottom right corner of the sheet.
    """
    
    row = last_row(sht)
    if row == 0: return None

    column = last_column(sht)

    r1c1_partial = "r1c1:r"+str(row)+"c"+str(column)
    r1c1_full = "='" + sht.name + "'!" + r1c1_partial
    

    return {'r1c1':{'full':r1c1_full,'partial':r1c1_partial},
            'ix':(row,column)}
    
def last_row(sht:xw.main.Sheet):
    """ Returns the row of the lowest non-empty cell in a sheet. """
    row_cell = sht.api.Cells.Find(What="*",
                   After=sht.api.Cells(1, 1),
                   LookAt=xw.constants.LookAt.xlPart,
                   LookIn=xw.constants.FindLookIn.xlFormulas,
                   SearchOrder=xw.constants.SearchOrder.xlByRows,
                SearchDirection=xw.constants.SearchDirection.xlPrevious,
                       MatchCase=False)
    
    if row_cell is None: return 0
    
    return row_cell.Row
        

def last_column(sht):
    """ Returns the row of the rightmost non-empty cell in a sheet. """
    column_cell = sht.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xw.constants.LookAt.xlPart,
                      LookIn=xw.constants.FindLookIn.xlFormulas,
                      SearchOrder=xw.constants.SearchOrder.xlByColumns,
                      SearchDirection=xw.constants.SearchDirection.xlPrevious,
                      MatchCase=False)
    
    c = column_cell.Column
    return c




if __name__ == "__main__":
    
    try:
        wb = xw.books[_book]
    except:
        wb = xw.Book(_book)
    
    wb.set_mock_caller()
    
    kill_broken_names()
    name_used_range_sheets_all()
