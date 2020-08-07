import os
from config import KEY_FOLDS_PATH


class PathReader(object):
    #  根据关键字folder，读取并yield里面的文件（暂时只支持doc）
    def __init__(self):
        self.key_word = ''
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
