import xlrd
#open the workbook
workbook = xlrd.open_workbook('page.xls')
worksheets = workbook.sheet_names()
#iterate through each row of each worksheet
for worksheet_name in worksheets:
    worksheet = workbook.sheet_by_name(worksheet_name)
    num_rows = worksheet.nrows -1
    curr_row = -1
    num_cells = worksheet.ncols-1
    #iterate through each row in each worksheet
    while curr_row<num_rows:
        curr_row+=1
        row = worksheet.row(curr_row)
        print 'Row:',curr_row
        curr_cell=-1
        #iterate through each cell on the row
        while curr_cell < num_cells:
            curr_cell+=1
            # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            print ' ', cell_type, ':', cell_value
    
    
    print worksheet_name

