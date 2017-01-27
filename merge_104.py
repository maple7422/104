import xlrd
import xlwt
import os

# 另存檔案的命名
new_save_file_name = '104_merge.xls'
# 儲存資料由第二行開始儲存 (第一行儲存資料抬頭)
new_rows = 1

file = xlwt.Workbook()
newtable= file.add_sheet('104',cell_overwrite_ok=True)

# 判斷需要建立的檔案是否存在，如果存在則刪除
filelist = os.listdir(os.getcwd())
for y in range(0,len(filelist)):
    if (filelist[y] == new_save_file_name):
        os.remove(new_save_file_name)
# 重新讀取當下目錄所有檔案
filelist = os.listdir(os.getcwd())
for y in range(0,len(filelist)):
    # 判斷 xls 才載入處理
    if (filelist[y].find('xls') > 0):
            print('正在處理 ' + filelist[y])
            data = xlrd.open_workbook(filelist[y])
            table = data.sheets()[0]
            nrows = table.nrows
            ncols = table.ncols
            # 處理第一行的資料抬頭
            for i in range(0,ncols):
                newtable.write(0,i,table.row_values(1)[i])
            # 每個檔案由第三列開始處理 (第一列是版權宣告，第二列是資料抬頭)
            for int_nrows in range(2,nrows):
                for i in range(0,ncols):
                    newtable.write(new_rows,i,table.row_values(int_nrows)[i])
                new_rows = new_rows + 1
print('建立 ' + new_save_file_name)
file.save(new_save_file_name)
print('完成')