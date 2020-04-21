import os
import sys
import fnmatch
import win32com.client
import codecs
import re
import rarfile


# zip_file = os.path.abspath(os.path.dirname(sys.argv[2]))
# zip_dir = r"C:\Users\86131\Documents\WeChat Files\a4528071\FileStorage\File\2020-04"
PATH = sys.argv[1]
# PATH = r"D:\work\data"
doc_path = os.path.join(PATH, "doc\\")
txt_path = os.path.join(PATH, "txt\\")
print(doc_path)
if not os.path.exists(doc_path):
    os.mkdir(doc_path)
if not os.path.exists(txt_path):
    os.mkdir(txt_path)


def unzip_rar(zip_dir, data_dir):
    for root, _, files in os.walk(zip_dir):
        for file in files:
            print(file)
            if fnmatch.fnmatch(file, "*.rar"):
                rf = rarfile.RarFile(file)
                rf.extractall(data_dir)


def convert_dir_to_txt():
    """
    将默认整个文件夹下的文件都进行转换
    :return:
    """
    for root, dirs, files in os.walk(doc_path):
        for _dir in dirs:
            pass
        print(root, dirs, files)
        for _file in files:
            if fnmatch.fnmatch(_file, '*.doc'):
                store_file = txt_path + _file[:-3] + 'txt'
            elif fnmatch.fnmatch(_file, '*.docx'):
                store_file = txt_path + _file[:-4] + 'txt'
            word_file = os.path.join(root, _file)
            dealer.Documents.Open(word_file)
            try:
                dealer.ActiveDocument.SaveAs(store_file, FileFormat=7, Encoding=65001)
            except Exception as e:
                print(e)
            dealer.ActiveDocument.Close()


def get_zh_ratio(string):
    count = 0
    length = len(string)
    for item in string:
        if 0x4E00 <= ord(item) <= 0x9FA5:
            count += 1
    return count / length



def convert_txt_to_sample():
    for root, _, files in os.walk(txt_path):
        wf = codecs.open(os.path.join(PATH, "output.txt"), encoding="utf-8", mode='w')
        for file in files:
            if fnmatch.fnmatch(file, "*.txt"):
                print(file)
                rf = codecs.open(os.path.join(root, file), encoding="utf-8")
                lines = []
                for line in rf.readlines():
                    #去掉空字符
                    pattern = re.compile(r"\s+")
                    line = re.sub(pattern, '', line)
                    #去掉序号
                    # pattern = re.compile(r"((\d+[.．、]\d{0,1})|(（{0,1}\d）))")
                    pattern = re.compile(r"(([\dA-z]+[.、．]\d{0,1})|([（(]{0,1}[\dA-z][）)].{0,1}))")
                    line = re.sub(pattern, '', line, count=1)
                    line = re.sub(r"^(.\d+)|^[!@#$%%.、，]{0,1}", "", line, count=1)
                    #删去长度小于20的句子
                    if len(line) < 20:
                        if len(line) == 0:
                            continue
                        elif get_zh_ratio(line) < 0.5 or len(line) < 15:
                            continue
                    else:
                        if get_zh_ratio(line) < 0.5:
                            continue
                    lines.append(line+"\n")
                rf.close()
                wf.writelines(lines)
                wf.write("\n")
        wf.close()


if __name__ == "__main__":
    dealer = win32com.client.gencache.EnsureDispatch('Word.Application')
    convert_dir_to_txt()
    convert_txt_to_sample()