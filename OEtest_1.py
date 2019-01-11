import numpy as np
import xlrd
import xlwt
import oe_tables

oe_0_0 = np.mat(
    [])


class oe():
    oe_table = oe_0_0
    col_num = 0
    state_num = 0
    row_num = 0


def chose_oe(input_dict):
    print(1)


def read_input():
    xlsfile = r"input.xlsx"  # 打开指 定路径中的xls文件
    book = xlrd.open_workbook(filename=xlsfile)  # 得到Excel文件的book对象，实例化对象
    sheet0 = book.sheet_by_index(0)  # 通过sheet索引获得sheet对象

    input_dict = {}
    i = 2
    try:
        while (sheet0.cell_value(i - 1, 0) != 'end'):

            key = sheet0.cell_value(i - 1, 1)
            # print(key)
            tmp_list = []
            j = 3
            try:
                while (sheet0.cell_value(i - 1, j - 1) != ''):
                    tmp_list.append(sheet0.cell_value(i - 1, j - 1))
                    j += 1
            except:
                print(j)
            input_dict[key] = tmp_list
            i += 1
    finally:
        print('input_dict', input_dict)
    return input_dict


# 去掉多餘的列
def col_reduce(oe_tmp, input_col_num):
    while (input_col_num < oe_tmp.col_num):
        oe_tmp.col_num -= 1
        print('before', oe_tmp)
        oe_tmp.oe_table = oe_tmp.oe_table[:, :oe_tmp.col_num]
        print('after', oe_tmp.oe_table)

    return oe_tmp



# 替換各因子的狀態，為後續行合併做準備
def state_replace(oe_tmp, input_dict):
    key_num = len(input_dict.keys())

    i = 0
    # 從因子開始替換
    while (i < key_num):
        j = 0
        # 從上到下逐行進行替換
        while (j < oe_tmp.row_num):
            a = oe_tmp.oe_table[j, i]
            # print(a)
            input_keys = list(input_dict.keys())
            # print(input_keys[0])
            max_index = len(input_dict[input_keys[i]]) - 1
            # 判斷正交表中的值是否超過了狀態數量
            if (a > max_index):
                oe_tmp.oe_table[j, i] = -1  # 設置為-1，後續處理
            # else:
            #     val=input_dict[input_keys[i]]
            #     val=val[a]
            #     oe_tmp.oe_table[j:i] = val
            j += 1
        i += 1
    return oe_tmp


def row_com(oe_tmp):
    i = 0
    # 从第一行开始比较
    while (i < oe_tmp.row_num):
        j = i + 1
        # 后面的每一行逐一对比
        while (j < oe_tmp.row_num):
            # 如果是已经合并了的行，直接跳过
            if (oe_tmp.oe_table[j, 0] == -2):
                j += 1
                continue
            k = 0
            # 对比行中的每一个值，-1排除
            while (k < oe_tmp.col_num):

                if (oe_tmp.oe_table[j, k] == -1):
                    k += 1
                    continue
                # 这个会不会有什么问题？？？
                # 前面的行中，如果有-1，而且后面的行中没有-1，则将后面的值赋值给前面的行
                elif (oe_tmp.oe_table[i, k] == -1):
                    oe_tmp.oe_table[i, k] = oe_tmp.oe_table[j, k]
                    k += 1
                    continue

                if (oe_tmp.oe_table[j, k] != oe_tmp.oe_table[i, k]):
                    break
                k += 1
            # 可以合并的行,-2进行标记
            if (k >= oe_tmp.col_num):
                oe_tmp.oe_table[j, :] = -2
            j += 1
        i += 1
    i = 0
    while (True):
        if (i >= oe_tmp.row_num):
            break
        if (oe_tmp.oe_table[i, 0] == -2):
            table_1 = oe_tmp.oe_table[0:i, :]
            table_2 = oe_tmp.oe_table[i + 1:, :]
            oe_tmp.oe_table = np.vstack((table_1, table_2))
            # 减少行数
            oe_tmp.row_num -= 1
            # 游标向上移动一个位置
            i -= 1
        if (i == 32):
            print(32)
            print(oe_tmp.oe_table)
        i += 1
    return oe_tmp


def output_xl(oe_tmp, input_dict):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    i = 0
    while (i < oe_tmp.col_num):
        j = 0
        while (j < oe_tmp.row_num):
            k = oe_tmp.oe_table[j, i]
            # -1则代表任意均可，手动填充
            if (k == -1):
                j += 1
                continue
            key = list(input_dict.keys())[i]
            val = input_dict[key][k]
            ws.write(j, i, val)
            j += 1
        i += 1

    wb.save('output.xls')
    print('output done')


if __name__ == '__main__':
    # 需要手动选择正交表
    flag = input('请确认input.xlsx已准备就绪，yes？')
    if (flag != 'yes' or flag != 'y'):
        exit(500)
    table_var = int(input('输入选择的正交表，因子数量:'))
    table_state = int(input('输入选择的正交表，状态数量:'))
    table_name = 'oe_' + str(table_state) + '_' + str(table_var)

    oe_tmp = oe()
    try:
        oe_tmp.oe_table = getattr(oe_tables, table_name)
        oe_tmp.row_num = oe_tmp.oe_table.shape[0]
        print('row_num:', oe_tmp.row_num)

        oe_tmp.col_num = table_var
        oe_tmp.state_num = table_state
    except:
        print('请检查选择的正交表是否存在')
        exit(404)

    input_dict = read_input()
    input_col = input_dict.keys()
    input_col_num = len(input_col)

    oe_tmp = col_reduce(oe_tmp, input_col_num)
    oe_tmp = state_replace(oe_tmp, input_dict)
    oe_tmp = row_com(oe_tmp)
    oe_tmp = row_com(oe_tmp)
    print('test:', oe_tmp.oe_table)
    output_xl(oe_tmp, input_dict)
    print('+++++++++++++++++++++++++++++++++++')
    print('注意：本程序预置的正交表均为标准正交表，交叉正交表需要去配置')
    print('已完成正交表整理、替换，请做以下操作')
    print('1.手动填充缺失的任意项，可以按照最常用的进行')
    print('2.删掉不可能的项')
    print('3.补充最常见、用户最喜欢的组合 ')
    print('+++++++++++++++++++++++++++++++++++')
