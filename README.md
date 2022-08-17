# Txt文本转Excel表格使用说明

### 运行环境

> 开发环境为python3.6.9，依赖包有：sys, os, re, openpyxl

在python3的环境下运行，需要额外安装`openpyxl`包，命令如下

```python
pip install openpyxl
```

### 启动

```shell
# 参数为文件
python3 txt2excel.py a_b_c_d.txt e_f_g_h.txt
# 参数为文件夹（里面只放txt）
python3 txt2excel.py folder_name1 folder_name2
```

+ txt按规定命名为` UUID_PROBE_TEMP_DATE.txt`
+ ~~保证`.py`与`.txt`处于同一目录~~
+ 保证此`output`目录下无`.xlsx`文件
+ 支持多文件以及多文件夹，只需保证路径准确

### 结果

程序会自动新建`output`文件夹，并将文件保存在此下面

所有txt文件中相同的`Sequence`将保存在同一个`"sequence_name".xlsx`文件中

#### 数据处理

在生成excel文件后，对表格数据进行处理

#### 表格合并

合并操作是在所有表格生成结束后再进行的，合并完成后再删除多余文件

### 参数
```python
mainfun = True  # 开启主功能
merger = True  # 开启合并功能
fun = True  # 开启转换功能
fun_key = ["VCCIO", "HVPP", "VLD"]  # 需要进行转换功能的excel文件前缀关键字
```