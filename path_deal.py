import os
from config import KEY_FOLDS_PATH


class PathReader(object):
    #  根据关键字folder，读取并yield里面的文件（暂时只支持doc）
    def __init__(self):
        self.key_word = ''
        self.file_name = ''
        try:
            self.key_folds = os.listdir(KEY_FOLDS_PATH)
        except FileNotFoundError:
            print('please keep the folder called docs_with_keyword')
            self.key_folds = []

    def read_key_folds_path(self):
        for key_fold in self.key_folds:
            self.key_word = key_fold
            fold_path = os.path.join(KEY_FOLDS_PATH, key_fold)
            if not os.path.isdir(fold_path):  # 跳过docs_with_keyword下非fold
                continue
            yield fold_path

    def get_key_word(self):
        return self.key_word

    def get_file_name(self):
        return self.file_name

    def get_word_path(self):
        for fold_path in self.read_key_folds_path():
            for file_name in os.listdir(fold_path):  # 这个路径是读取出来的，不考虑有问题
                if file_name.endswith('.doc') and not file_name.startswith('~$'):  # 由于目前只处理word，所以在此判断
                    if os.path.exists(os.path.join(fold_path, file_name + 'x')):
                        print(file_name, '已经转换过doc，跳过本次转换')
                        continue
                    self.file_name = file_name
                    file_path = os.path.join(fold_path, file_name)
                    yield file_path
                elif file_name.endswith('.docx') and not file_name.startswith('~$'):
                    self.file_name = file_name
                    file_path = os.path.join(fold_path, file_name)
                    yield file_path
