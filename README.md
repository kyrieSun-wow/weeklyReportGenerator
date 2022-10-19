# weeklyReportGenerator
Generate weekly team reports in batches based on DingDing's export files


# Foreword：


## Precondition:
### 1.windows 系统需要支持python。
    使用前请查看windows系统是否支持python，  可在CMD中敲入python 命令。
    如需安装则可用本程序自带的python安装包，也可以从官网下载更新的windows installer:   https://www.python.org/downloads/release/python-3107/
    注：安装python的时候，一定要勾选   Add Python 3.9 toPATH    选项
    windows安装python可参考： https://blog.csdn.net/m0_63604019/article/details/124203850

### 2.windows 系统需要支持python第三方库 openpyxl。
    安装方式： cmd 下敲入pip install openpyxl      （安装过程大概要1分钟）
 
### 3.在TeamMateConfig.txt 文件中查看，添加所需的team, 以及成员的中英文名对照字典。

### 4.使用前请关闭将要接收数据的xlsx表格，不然有可能导致数据写入失败。

# How to use：
打开weeklyReportGenerator.bat 文件，按指示操作即可。


