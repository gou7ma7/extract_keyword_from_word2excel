import os

from win32com import client as wc
from docx import Document

from config import KEY_FOLDS_PATH


def read_key_folds(path):
    try:
        listdir = os.listdir(path)
    except FileNotFoundError:
        print('please keep the folder called docs_with_keyword')
        return []

    for fold_name in listdir:
        global key_word
        key_word = fold_name

        fold_path = os.path.join(path, fold_name)
        if not os.path.isdir(fold_path):
            continue
        yield fold_path


def deal_doc_folds():
    for fold_path in read_key_folds(KEY_FOLDS_PATH):
        for file in os.listdir(fold_path):  # 这个路径是读取出来的，不考虑有问题
            doc_path = os.path.abspath(os.path.join(fold_path, file))
            if file.endswith('.doc'):
                docx_path = doc_path + 'x'
                if os.path.exists(docx_path):
                    continue
                print('正在转换为docx格式', doc_path)
                doc2docx(doc_path)
            elif file.endswith('.docx'):
                ex_para(doc_path, key_word)
        print("处理关键字文件夹：", fold_path, key_word)


def ex_para(doc_path, kw):
    document = Document(doc_path)
    for paragraph in document.paragraphs:
        if kw in paragraph.text:
            print(paragraph.text)


def doc2docx(doc_path):
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)  # 目标路径下的文件
    doc.SaveAs(doc_path + 'x', 12, False, "", True, "", False, False, False,
               False)  # 转化后路径下的文件
    doc.Close()
    word.Quit()


def get_info():
    pass


if __name__ == '__main__':
    key_word = ''
    deal_doc_folds()
    get_info()
    print()

    print('ok')
