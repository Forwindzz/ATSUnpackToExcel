
# Unpack Analysis scripts

This scripts only used for unpack and export ATS Data
Currently do not support English, only support Chinese, but you can still use it. I will add English version if anyone want this.

分析脚本，会生成表格，暂时不支持英语

代码很乱，能跑就行的那种

Support ATS 1.4 only

# 环境要求

- Python 3
- openpyxl, json, yaml

你可以通过
```shell
pip install openpyxl json yaml
```
来安装这些库
# 构建

创建`all_2` `output` `vis_output`文件夹在仓库根目录中
使用[AssetRipper](https://github.com/AssetRipper/AssetRipper)导出到`all_2`文件夹
你的`all_2`文件夹下面应该只有`AuxiliaryFiles`和`ExportedProject`文件夹

然后在根目录打开powershell/cmd/windows terminal运行：
```bat
./parse.bat
```
里面会执行一堆python脚本

最终输出在`output`和`output_vis`
- `output` 将数据整理成json打包
- `output_vis` 将数据转化为可视化表格
- `uuid_mapping.json` `uuid_path_mapping.json` prefab的uuid mapping
- `output/display_names.json` 将所有翻译和语言信息集合起来

# 配置文件

- `add_translate.json` 额外语言表，能凑合翻译就行
- `test_average_settings.json5` 配置期望计算，会对一些特殊空地组进行额外计算
- `translator_index.json5` 给空地中不同组别的生成配置命名
- `*.xlsx` 这个是表格的第一页，我通常会写一些说明在里面

# 脚本

- `gather_*.py` 用来收集特定类型的数据
- `replace_text_uuid.py` 将`text.txt`文件中uuid全部换成prefab文件名，方便阅读
- `gen_sheet_glades.py` 生成excel表格，会使用`output`文件夹下的数据

