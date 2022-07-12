import pandas as pd
import os
import sys

def list_all_excel_file(path, all_files = []):    
    if os.path.exists(path):
        files = os.listdir(path)
    else:
        print('this path not exist')
    for file in files:
        if os.path.isdir(os.path.join(path,file)):
            list_all_excel_file(os.path.join(path,file), all_files)
        else:
            if (file.endswith('.xls') or file.endswith('.xlsx')):
                all_files.append(os.path.join(path, file))
    return all_files


def search_keyword(value, key):
    for fileName in files:
        items = pd.read_excel(fileName)

        for index, row in items.iterrows():
            if ((key in row.keys()) and (str(row[key]) == value)):
                print ('找到了[', key, '为', value, ']的人,' , 'TA在这个文件[', fileName, ']' '的第[', index + 1, ']行！现在尝试打开文件...' )
                os.system("explorer " + fileName)

if __name__ == '__main__':
    files = list_all_excel_file(r'.')
    search_keyword(sys.argv[2], sys.argv[1])