
# -PDZ-PDF-

超星阅读器PDZ文件转PDF工具

以管理员身份打开cmd终端
在终端直接输入以下指令
python py文件地址 “PDZ文件地址”
例如：
python D:\ssreader\pdz2pdf.py "D:\ssreader\xxx.pdz"
回车即可运行
自动识别页数可能最后会少几页
也可以选择手动输入
还是同一个例子
python D:\ssreader\pdz2pdf.py "D:\ssreader\xxxxx.pdz" --pages 50
最后一页放在第三页左右的位置是个小bug，无伤大雅
