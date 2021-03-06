# ExcelDiffTool
Excel表格对比工具，pyQT可视化界面，批量对比。
![image](https://github.com/smilefacehh/ExcelDiffTool/blob/main/screen.png)

### 目录
- [安装依赖](#t1)
- [使用方法](#t2)
- [代码说明](#t3)
- [文件说明](#t4)

### <span id="t1">1.安装依赖</span>
```
1) 安装 python3, pip3
2) 安装依赖项 pip3 install pyqt5 xlrd xlsxwriter
```

### <span id="t2">2.使用方法</span>
使用pyinstaller打包py脚本，生成可执行exe文件

```pyinstaller.exe -F -w -i emma2.ico .\MainWindow.py .\excel_diff.py .\images.py```

运行程序，exe所在目录会生成日志文件`log.txt`，以及表格对比结果文件`xx_diff.xlsx`

颜色说明：
1. 黄色：修改
2. 蓝色：新增
3. 红色：删除
4. 绿色：行位置有变化

### <span id="t3">3.代码说明</span></span>
Excel表格对比结果的准确性，依赖于两个参数，可以调整这两个参数来尝试取得更好的对比结果
```
# 行距阈值。计算匹配行时，选择相邻100行作为候选匹配
COMPARE_LINE_DIS = 100

# 内容-行距权重。计算行相似度，根据单元格内容匹配程度，与行位置差异进行权重相加，计算相似度，0.8表示内容匹配的权重
COMPARE_GAMMA = 0.8
```

### <span id="t4">4.文件说明</span>
+ MainWindow.ui 用Qt Designer设计并生成的图形界面文件
+ MainWindow.py 程序入口，使用`pyuic5 -o MainWindow.py MainWindow.ui`生成基本框架代码
+ excel_diff.py 表格对比逻辑
+ images.qrc    用于将ico生成py
+ images.py     ico对应py，用于设置icon，使用`pyrcc5 images.qrc -o images.py`生成
