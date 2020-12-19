Excel转换器
=========
中文
* 编写及调试环境
  Microsoft Visual Studio 2019 Community
* 基础运行(Debug)
  * Windows 10 Enterprise x64
---------
* 目的  
  Python提供了xlrd库支持在没有安装Excel的环境下访问、读取、写入xls/xlsx文件，这为跨环境文件处理提供了便利。  
于是我们开发此程序通过Python的这个优点读取、保存xls文件为csv文件，后者对大多数只需要关注内容的程序来说更为友好。  
  程序将读取所有表，并保存为一系列文件以存储所有信息
* 使用方法   
  * eph [-h|-c|-p|--help|--csv|--print] {filename}
  * -h --help ： 打印帮助菜单
  * -c --csv : 生成csv文件。保留这个指令的原因是后续可能会开发其他格式和协议文件
* 运行方式  
  * 必须使用命令行按照指令操作，这为程序之间互相调用提供了便利

Excel files converter
=========
English
* Developing and debugging environment
   Microsoft Visual Studio 2019 Community
* Basic operation (Debug)
   * Windows 10 Enterprise x64
---------
* Purpose  
Python provides the xlrd library to support accessing, reading, and writing xls/xlsx files in an environment where Excel is not installed, which provides convenience for cross-environment file processing.  
So we developed this program to read and save xls files as csv files through this advantage of Python. The latter is more friendly to most programs that only need to focus on content.  
The program will read all tables and save them as a series of files to store all information
* Instructions  
   * eph [-h|-c|-p|--help|--csv|--print] {filename}
   * -h --help: Print the help menu
   * -c --csv: Generate csv file. The reason for retaining this instruction is that other formats and protocol documents may be developed in the future
* Operation mode
   * You must use the command prompt to follow the instructions, which provides convenience for calling each other between programs
