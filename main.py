
from excel_deal import ExcelSaver
from path_deal import PathReader
from word_dealer import WordDealer

if __name__ == '__main__':
    pr = PathReader()

    es = ExcelSaver()
    for word_path in pr.get_word_path():
        key_word = pr.get_key_word()  # 因为本业务中关键字是从文件夹的名字里面读取的来的，所以是PathReader的方法得到
        file_name = pr.get_file_name()
        wd = WordDealer(word_path, key_word)

        completion_year, province, para = wd.extract_paragraph()
        print(completion_year, province, para)
        es.add_row(file_name, completion_year, province, key_word, para)  # 目前excel格式也是固定的
    es.save_xlsx()

    print('ok')
