# extract_paragraphs_in_word_by_keywords

## Why
1. .doc格式的文件是二进制文件， .docx的文件是xml，两者完全不同。python要处理需要pywin32转换，要么antiword；
2. antiword是linux平台的，考虑到最后程序要在win上面跑，所以先做转换；