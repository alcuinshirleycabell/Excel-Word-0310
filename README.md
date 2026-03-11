# ExcelToWord 使用指南

## 功能介绍

Excel数据批量生成Word报告，支持：
- ✅ 智能匹配占位符（简写自动匹配全称）
- ✅ 本地GUI运行 / GitHub Actions云端运行 / Windows EXE
- ✅ 灵活的文件命名规则

## 三种使用方式

### 方式1：本地运行（Python）
```bash
pip install python-docx pandas openpyxl
python batch_word_gen.py
```

### 方式2：GitHub Actions（云端）
1. 上传模板和Excel到仓库，或提供下载链接
2. 进入Actions → "Excel to Word Batch Generator" → Run workflow
3. 下载生成的文件

### 方式3：Windows EXE（双击运行）
1. 进入Actions → "Build Windows EXE" → Run workflow
2. 等待构建完成，下载exe文件
3. 双击运行，选择文件生成

## 占位符规则

Word模板中支持：
- `客户姓名` - 纯文本（推荐）
- `__客户姓名__` - 双下划线
- `{客户姓名}` - 花括号

智能匹配示例：
- `日均总资产` → `日均总资产(近一年)`
- `盈亏金额` → `盈亏金额(近一年)`

## 文件命名

- 按列命名：输入 `客户姓名`
- 自定义模板：`报告_{客户姓名}_{序号}`

变量：`{列名}`、`{序号}`
