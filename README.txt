androidstring2po说明

扫描Android代码中，统计国际字符串到totol.xlsm中

安装要求:
需要安装以下Python模块，当然他们大部分都能自动安装，请参考各自的安装说明
babel >= 1.0dev (only the dev version has support for contexts)
http://babel.edgewall.org/
lxml
http://codespeak.net/lxml/
argparse
http://argparse.googlecode.com/
a2po
https://github.com/huzhennan/android2po

使用说明：
1) cd到Android项目根目录
2）python utils.py init [language] 初始化生成totol.xlsm统计表格
此处：language：中文对应 zh_CN, 德文对应 de.
3)更新字符串
4)python utils.py import [language] 从totol.xlsm导入修改后数据到各个模块
