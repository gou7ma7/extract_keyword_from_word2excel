# extract_keyword_paragraphs

## setup
1. https://www.python.org/getit/ 官网下载python;
2. 安装且配置环境变量等ok之后，在CMD中cd 到本项目根目录，pip install pipenv;
3. 启动项目环境pipenv install，网不好可能需要多次重复；
4. 创建一个关键字文件夹，里面装欲提取的word，再将关键字文件夹等全部放到根目录下面的docs_with_keyword文件夹下
5. python main.py 启动项目。

## what
1. 纯业务的项目，受朋友之托，在很多个word中提取含有特定关键字的段落和一些其他信息，汇总整理成excel；
2. word主要格式有.doc和.docx，前者是旧的格式储存方式为二进制文件，python处理的时候需要先使用win32com库转换成储存方式为xml的.doc处理；
（也可以使用antiword库直接解析，但该库是linux平台的，考虑到最后程序要在win上面跑，遂放弃）
3. 目前主要分为三个模块，在main.py里面通过路径获得文件，处理完之后汇总到excel，这样处理方便扩展模块，以后可以接其他的文件类型；
4. PathReader，主要是处理一些路径不存在，文件夹重复的问题；
5. WordDealer，完成.doc -> .docx，解析，使用python-docx库；
6. ExcelSaver，汇总并储存为.xlsx，使用xlsxwriter库，是一个专注于写入表格的库。

## 可能的坑
1. WordDealer里面win32com打开.doc或者储存为.docx时候会有千奇百怪的问题，查询发现也是历史遗留，好在可以通过retry暂时解决。