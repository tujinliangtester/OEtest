import xlrd

def readXl(xlsfile,row,col):
    book = xlrd.open_workbook(filename=xlsfile)#得到Excel文件的book对象，实例化对象
    sheet0 = book.sheet_by_index(0) # 通过sheet索引获得sheet对象

    i=row
    try:
        while(sheet0.cell_value(i,0)!=''):
            tmp_list=[]
            j=col
            try:
                while(sheet0.cell_value(i,j)!=''):
                    tmp_list.append((int)(sheet0.cell_value(i,j)))
                    j+=1
            except:
                j=0

            print(tmp_list)
            i+=1
    finally:
        print('input_dict')

if __name__=='__main__':
    readXl('oe.xlsx',0,0)