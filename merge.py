import xlrd, xlwt
from pathlib import Path

def merge(srcpath, pattern, destpath):
    i = 0
    dest_book = xlwt.Workbook()
    dest_sheet = dest_book.add_sheet('Sheet1')
    for path in Path(srcpath).glob(pattern):
        src_book = xlrd.open_workbook(str(path))
        src_sheet = src_book.sheet_by_index(0)
        for row in src_sheet.get_rows():
            j = 0
            for cell in row:
                dest_sheet.write(i, j, cell.value)
                j += 1
            i += 1

    dest_book.save(destpath)
    

if __name__ == '__main__':
    merge('data', '*.xlsx', 'result/merge.xls')
    
    
