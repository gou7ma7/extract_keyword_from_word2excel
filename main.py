import os

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


def read_docs():
    for fold_path in read_key_folds(KEY_FOLDS_PATH):
        for file in os.listdir(fold_path):  # 这个路径是读取出来的，不考虑有问题
            if not (file.endswith('.doc') or file.endswith('.docx')):
                continue
            doc_path = os.path.abspath(os.path.join(fold_path, file))
            # try:
            #     document = Document(doc_path)
            #     print('document', document)
            #
            # except Exception as e:
            #     print(e)
            break
        print("fold_path", fold_path, key_word)
    return []


if __name__ == '__main__':
    key_word = ''
    for doc in read_docs():
        print('doc', doc)
    print()

    KEY_FOLDS_PATH = './empty_test'
    for doc in read_docs():
        print('doc', doc)
    print()

    print('ok')
