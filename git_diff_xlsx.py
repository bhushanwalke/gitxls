
import xlrd as xl
import sys
import unicodedata

def parse(infile,outfile):
    """
    Converts an Excel file into text
    Returns a formatted text file for comparison using git diff.
    """

    book = xl.open_workbook(infile)

    num_sheets = book.nsheets

    print book.sheet_names()

#   print "File last edited by " + book.user_name + "\n"
    outfile.write("File last edited by " + book.user_name + "\n")

    def get_cells(sheet, rowx, colx):
        try:
            value = unicode(sheet.cell_value(rowx, colx))
        except:
            value = ''
        if (rowx,colx) in sheet.cell_note_map.keys():
            value += ' <<' + unicodedata.normalize('NFKD', unicode(sheet.cell_note_map[rowx,colx].text)).encode('ascii', 'ignore') + '>>'
        return value

    # loop over worksheets

    for index in range(0,num_sheets):
        # find non empty cells
        sheet = book.sheet_by_index(index)
        outfile.write("=================================\n")
        sheetname = unicodedata.normalize('NFKD', unicode(sheet.name)).encode('ascii', 'ignore')
        outfile.write("Sheet: " + sheetname + "[ " + str(sheet.nrows) + " , " + str(sheet.ncols) + " ]\n")
        outfile.write("=================================\n")
        for row in range(0,sheet.nrows):
            for col in range(0,sheet.ncols):
                content = get_cells(sheet, row, col)
                if content <> "":
                    output = '    ' + xl.cellname(row,col) + ':\n        '
                    output += unicodedata.normalize('NFKD', unicode(content)).encode('ascii', 'ignore')
                    output += '\n'
                    outfile.write(output)
        print "\n"

# output cell address and contents of cell
def main():
    args = sys.argv[1:]
    if len(args) != 1:
        print 'usage: python git_diff_xlsx.py infile.xlsx'
        sys.exit(-1)
    outfile = sys.stdout
    parse(args[0],outfile)

if __name__ == '__main__':
    main()
