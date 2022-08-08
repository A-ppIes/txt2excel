# Txt文本转Excel表格使用说明

### 运行环境

> 开发环境为python3.6.9，依赖包有：sys, os, re, openpyxl

在python3的环境下运行，需要额外安装`openpyxl`包，命令如下

```python
pip install openpyxl
```

### 启动

```shell
python3 txt2excel.py a_b_c_d.txt e_f_g_h.txt
```

+ txt按规定命名为` UUID_PROBE_TEMP_DATE.txt`
+ 保证`.py`与`.txt`处于同一目录
+ 保证此`.py`目录下无`.xlsx`文件

### 结果

所有txt文件中相同的`Sequence`将保存在同一个`"sequence_name".xlsx`文件中

#### 数据处理

在生成excel文件后，对表格数据进行处理

#### 表格合并

合并操作是在所有表格生成结束后再进行的，合并完成后再删除多余文件