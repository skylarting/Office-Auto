# Office-Auto

Windows Excel 单元格批量汇总工具。

## 功能

- 选择一个 `.xlsx` / `.xlsm` 文件；
- 选择文件夹并批量处理其中全部工作簿；
- 自定义需要提取的单元格地址；
- 支持逗号、空格、换行和中英文分号分隔；
- 保留真实数值，并复制源单元格数字显示格式；
- 可新建独立汇总文件；
- 可在每个源工作簿最前面插入或替换“汇总”工作表；
- 文件夹生成独立汇总文件时，自动增加“文件名称”和“Sheet名称”列；
- 自动跳过临时文件及已经生成的汇总文件；
- 支持 `.xlsm`，写回源文件时保留 VBA 项目。

## EXE 使用方法

1. 从 GitHub Actions 下载并解压 `ExcelSummary-Windows`；
2. 双击 `ExcelSummary.exe`；
3. 选择一个 Excel 文件或包含多个 Excel 文件的文件夹；
4. 填写需要提取的单元格，例如：

   ```text
   I8, J8, L8, I17, J17, L17
   ```

5. 选择输出方式：

   - **新建一个汇总文件**：单文件生成一个汇总文件；文件夹模式将所有工作簿合并到一个汇总文件；
   - **在源工作簿最前面插入汇总表**：逐个修改源工作簿。

6. 点击“开始提取”。

> 写入源工作簿会修改原文件，建议提前备份。

## 公式单元格

程序读取 Excel 文件中上次保存的公式计算结果。如果公式结果为空，请先用
Excel 或 WPS 打开源文件、完成计算并保存，再运行本工具。

## 本地运行

```bash
python -m pip install openpyxl
python export_summary.py
```

命令行模式仍然可用：

```bash
python export_summary.py "测试数据.xlsx" \
  -o "汇总结果.xlsx" \
  -c "I8,J8,L8"
```

## 测试

```bash
python -m unittest -v
```

## 构建 Windows EXE

推送到 GitHub 后，`Build Windows EXE` 工作流会在 Windows 环境中运行测试并
生成 `ExcelSummary.exe`。在 Actions 运行页面底部下载
`ExcelSummary-Windows` 构建产物即可。
