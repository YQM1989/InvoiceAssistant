# 发票助手（Windows 本地版）

## 功能
- 导入单个/批量 PDF 发票
- 自动识别并重命名：`公司名_发票号.pdf`
- 重复发票拦截（公司名+发票号）
- 记忆功能：SQLite 持久化
- 按公司模糊搜索、按月份筛选（支持双重筛选）
- 分页展示（每页 20 条）
- 编辑/删除发票，编辑后自动同步文件名
- 导出筛选结果（PDF + 记账 CSV）
- 分目录导出（按 `月份/公司` 自动建目录）
- 供应商白名单（手工添加 / Excel 导入）
- 对账功能（公司名 + 发票号 + 金额）
- 高金额（>10万）导入前二次确认
- OCR 增强识别（文本提取不足时自动启用 RapidOCR）

## 运行方式
双击 `run.bat`

或者命令行运行：

```powershell
.\.venv\Scripts\python.exe app.py
```

## 文件说明
- `app.py`：主界面和业务逻辑
- `storage.py`：数据库与持久化
- `invoice_parser.py`：PDF 解析与字段识别
- `build_exe.bat`：一键打包 EXE
- `invoice_assistant.db`：程序首次运行后自动创建
- `invoice_files/`：导入后的发票文件归档目录
- `exports/`：导出文件默认目录（可自行选择其他目录）
- `dist/InvoiceAssistant/InvoiceAssistant.exe`：打包后可执行文件

## 打包 EXE
双击 `build_exe.bat`，打包完成后在 `dist/InvoiceAssistant/` 目录获取 `InvoiceAssistant.exe`。

## OCR 说明
- 程序优先使用 PDF 文本提取。
- 当识别字段不足时，会自动启用 OCR（扫描件识别能力更强）。
- OCR 仍可能有误差，可在列表中手动编辑修正。
