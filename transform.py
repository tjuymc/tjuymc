import os
import os.path
import shutil
import win32com.client as win32

def transform():
    ## 根目录
    rootdir = input('Enter filedir:')
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(rootdir):
        for fn in filenames:
            filedir = os.path.join(parent, fn)
            print(filedir)

            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56,
            # 后缀名的大小写不通配，需按实际修改：xls，或XLS
            wb.SaveAs(filedir.replace('XLS', 'xlsx'), FileFormat=51)  
            wb.Close()
            excel.Application.Quit()

    source_path = os.path.abspath(rootdir)
    target_path = os.path.abspath(input('Enter target dir:'))
    if not os.path.exists(target_path):
        os.makedirs(target_path)
    if os.path.exists(source_path):
        # root 所指的是当前正在遍历的这个文件夹的本身的地址
        # dirs 是一个 list，内容是该文件夹中所有的目录的名字(不包括子目录)
        # files 同样是 list, 内容是该文件夹中所有的文件(不包括子目录)
        for root, dirs, files in os.walk(source_path):
            for file in files:
                src_file = os.path.join(root, file)
                if src_file.endswith('.xlsx'):
                   shutil.copy(src_file, target_path)
                   print(src_file)




