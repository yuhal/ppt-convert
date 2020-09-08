# 感谢

- [python-pptx](https://github.com/scanny/python-pptx "python-pptx")

# 简介

> 一些 Python 操作 PPT 的案例，可以用于对 PPT 进行批量转换为 PDF、PNG，还可以抓取 PPT 中的文本信息。

# 目录

```
├─catch_text.py
# pptx 抓取text
├─catch_format_text.py
# pptx抓取text并按段落格式
├─convert_png.py
# pptx转换png
├─convert_pdf.py
# pptx转换pdf
├─convert_pptx.py
# ppt转换pptx
├─convert_split_pdf.py
# pptx转换并拆分pdf
├─sample.pptx
# pptx文件样例
├─sample.ppt
# ppt文件样例
```

| 功能  | 运行系统  |
| ------------ | ------------ |
| pptx 抓取 text  | 任何系统 |
|  pptx 转换 png、pdf <br/>ppt 转换 pptx<br/>pptx 转换并拆分 pdf<br/>  | Windows（需安装 office）  |

# 启动

- 下载

```
$ git clone https://github.com/yuhal/ppt-convert.git
```

- 安装 python-pptx

```
$ pip install python-pptx
```

- 执行

```
$ python catch_format_text.py
{0: ['乘坐时光机 |', '回到那些年 |', '我之所以到现在还怎么没用，是因为我不想离开哆啦A梦'], 1: ['你看，不倒翁站起来了，大雄也可以自己站起来啊!'], 2: ['你受伤的时候，我永远都在。']}
```

# License 

[MIT](https://github.com/yuhal/ppt-convert/blob/master/LICENSE "MIT")