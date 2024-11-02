
# Unpack Analysis scripts

Support ATS 1.4 only
只支持1.4版本

Analyze and export json data and excel readable sheet
分析并导出json和表格数据

# Environment require 

- Python 3
- openpyxl, json, yaml

Use this command to install libraries:
```shell
pip install openpyxl json yaml
```

# Build

English Version:
Create folders:`all_2` `output` `vis_output` in the root
Use [AssetRipper](https://github.com/AssetRipper/AssetRipper), export game files to `all_2` folder
Your `all_2` folder should only contain`AuxiliaryFiles` and `ExportedProject`folder

open powershell/cmd/windows terminal and run this in the root：
```bat
./parse.bat
```
It will run many python scripts

The final output is to `output` and `output_vis`
- `output` packed data in json format
- `output_vis` readable excel sheets
- `uuid_mapping.json` `uuid_path_mapping.json` prefab's uuid mapping
- `output/display_names.json` Localization file

# Config file

- `settings/lang.json` Change current Language
- `add_translate.json` For translation
- `test_average_settings.json5` Expectation calculation for glades
- `translator_index.json5` name the spawn group 
- `template_*.xlsx` the first page of the excel sheet, there are some descriptions here

# Scipts

- `gather_*.py` gather data and output it to `*.json`
- `replace_text_uuid.py` replace `text.txt` uuid string with prefab filename for higher readability
- `gen_sheet_glades.py` generate excel sheet，use the data in `output` and output to `output_vis`

# 构建

Chinese Version

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

- `settings/lang.json` 设置当前语言
- `add_translate.json` 额外语言表，能凑合翻译就行
- `test_average_settings.json5` 配置期望计算，会对一些特殊空地组进行额外计算
- `translator_index.json5` 给空地中不同组别的生成配置命名
- `template_*.xlsx` 这个是表格的第一页，我通常会写一些说明在里面

# 脚本

- `gather_*.py` 用来收集特定类型的数据
- `replace_text_uuid.py` 将`text.txt`文件中uuid全部换成prefab文件名，方便阅读
- `gen_sheet_glades.py` 生成excel表格，会使用`output`文件夹下的数据，文件会最终输出到`output_vis`

