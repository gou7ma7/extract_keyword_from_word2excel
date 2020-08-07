# extract_paragraphs_in_word_by_keywords

## setup
1. https://www.python.org/getit/ 官网下载python;
2. 安装且配置环境变量等ok之后，在CMD中cd 到本项目根目录，pip install pipenv;
3. 启动项目环境pipenv install，网不好可能需要多次重复；
4. python main.py 启动项目。

## Why
1. .doc格式的文件是二进制文件， .docx的文件是xml，两者完全不同。python要处理需要pywin32转换，要么antiword；
2. antiword是linux平台的，考虑到最后程序要在win上面跑，所以先做转换；