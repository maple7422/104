import xlrd
import xlwt
import os

new_save_file_name = '104_merge.xls'
new_rows = 1

file = xlwt.Workbook()
newtable= file.add_sheet('104',cell_overwrite_ok=True)

filelist = os.listdir(os.getcwd())
for y in range(0,len(filelist)):
    if (filelist[y] == new_save_file_name):
        os.remove(new_save_file_name)
filelist = os.listdir(os.getcwd())
for y in range(0,len(filelist)):
    if (filelist[y].find('xls') > 0):
            print('正在處理 ' + filelist[y])
            data = xlrd.open_workbook(filelist[y])
            table = data.sheets()[0]
            nrows = table.nrows
            ncols = table.ncols

            for i in range(0,ncols):
                newtable.write(0,i,table.row_values(1)[i])
            for int_nrows in range(2,nrows):
                for i in range(0,ncols):
                    newtable.write(new_rows,i,table.row_values(int_nrows)[i])
                new_rows = new_rows + 1
print('建立 ' + new_save_file_name)
file.save(new_save_file_name)
print('完成')






