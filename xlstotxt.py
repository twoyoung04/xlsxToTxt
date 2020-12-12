
#-*- coding:UTF-8 -*-
import xlrd
import os
def strs(row):
    """
    :返回一行数据
    """
    try:
        values = "";
        for i in range(len(row)):
            if i == len(row) - 1:
                values = values + str(row[i]) + "\n"
            else:
                #使用“，”逗号作为分隔符
                values = values + str(row[i]) + "\t" 
        return values
    except:
        print('strs error')
        raise

def check_xlsx(file_path):
    try:
        ext = file_path.rsplit('.')[1]
        return ext == 'xlsx' or ext == 'xls'
    except:
        print('file has no extention\n')
    return false

def xls_to_txt(file_path, output_path = None):
    """
    :excel文件转换为txt文件
    :param xls_name excel 文件名称
    :param txt_name txt   文件名称
    """

    if not check_xlsx(file_path):
        return

    file_name = file_path.rsplit('.')[0]

    output_ext = 'txt'
    output_path = output_path if output_path else file_name + '.' + output_ext
    print(file_path, output_path)

    try:
        data = xlrd.open_workbook(file_path)
    except:
        print('open xlsx error')
        return
    
    try:
        FOUT = open(output_path, "w+") 
        table = data.sheets()[0] # 表头
        nrows = table.nrows  # 行数
        #如果不需跳过表头，则将下一行中1改为0
        for ronum in range(0, nrows):
            row = table.row_values(ronum)
            values = strs(row) # 条用函数，将行数据拼接成字符串
            FOUT.writelines(values) #将字符串写入新文件
            print(values)
        FOUT.close() # 关闭写入的文件
    except:
        print('xls_to_txt error')
        pass

def xls_to_txt_dir(dir_path):
    path_list = os.listdir(dir_path)
    for file_path in path_list:
        if check_xlsx(file_path):
            xls_to_txt(file_path)


if __name__ == '__main__':
    
    dir_path = './'
    xls_to_txt_dir(dir_path)