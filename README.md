# 文件转换器

根据 excel 文件，替换 office word 文件的指定文本，自动生成文件。
替换规则:
- office word 文件中包含需要替换的文本。例如: %%name%%, %%cellphone%%
- excel 文件中表头文本用 / 分隔，用 / 分隔的最后一组文本文本，会被识别并替换。例如: 名字/%%name%%
- 输出文件位于主程序目录下

# 命令行用例
```
$python pattern_relpacer.py -x resources\source.xlsx info -w resources\template.docx
```

打印帮助:
```
$python pattern_relpacer.py -h
usage: pattern_relpacer.py [-h] [-x EXCEL_FILE EXCEL_FILE]
                           [-w TEMPLATE_WORD_FILE]

optional arguments:
  -h, --help            show this help message and exit
  -x EXCEL_FILE EXCEL_FILE
                        input excel file
  -w TEMPLATE_WORD_FILE
                        input temple word file
```

# 图形工具
在 release 页面下载文件，解压点击 e2w.exe 即可使用
