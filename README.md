# excel2json

将 Excel 文件（`.xlsx` / `.xls`）转换为 JSON 的命令行工具，支持单文件转换和多文件合并。

**特性：**
- 字段顺序固定，按 Excel 表头顺序输出
- 支持 NDJSON 和 JSON 数组两种输出格式
- 多文件合并时自动补全缺失字段

## 安装

```bash
go install github.com/muzimu/excel2json@latest
```

或克隆仓库后本地安装：

```bash
git clone https://github.com/muzimu/excel2json
cd excel2json
go install .
```

## 用法

```
excel2json <input> <output.json> [flags]
```

### 参数

| 参数 | 简写 | 说明 |
|------|------|------|
| `--sheet` | `-s` | 指定读取的工作表名称，默认读取第一个表 |
| `--list` | `-l` | 输出 JSON 数组格式，默认输出 NDJSON（每行一个对象） |
| `--pretty` | `-p` | 格式化输出，仅对 `--list` 有效 |
| `--no-header` | `-n` | 不将第一行作为字段名，改用列索引（`col0`, `col1`, ...） |

### 示例

```bash
# 基本用法，默认输出 NDJSON
excel2json input.xlsx output.json

# 读取指定工作表
excel2json input.xlsx output.json -s "Sheet2"

# 输出 JSON 数组
excel2json input.xlsx output.json -l

# 输出格式化的 JSON 数组
excel2json input.xlsx output.json -l -p

# 不使用表头，列名改为 col0, col1, ...
excel2json input.xls output.json -n
```

### 字段顺序

输出 JSON 时字段顺序固定：

- **有表头**：按 Excel 表头顺序输出
- **无表头**：按 `col0`, `col1`, `col2`, ... 顺序输出

## 合并多个文件

```
excel2json merge [flags] <file1.xlsx> <file2.xlsx> ...
```

### merge 参数

| 参数 | 简写 | 说明 |
|------|------|------|
| `--output` | `-o` | 输出文件路径，默认 `merge.json` |
| `--sheet` | `-s` | 指定读取的工作表名称，默认读取第一个表 |
| `--list` | `-l` | 输出 JSON 数组格式，默认输出 NDJSON |
| `--pretty` | `-p` | 格式化输出，仅对 `--list` 有效 |
| `--no-header` | `-n` | 不将第一行作为字段名，改用列索引 |
| `--no-source` | - | 关闭来源标记（默认添加 `_source` 字段） |

### 合并规则

1. **表头取并集**：所有文件的列名合并，缺失字段填空字符串
2. **字段顺序固定**：按 Excel 表头顺序输出，`_source` 字段在最后
3. **来源标记**：默认添加 `_source` 字段，记录数据来源文件名（不含扩展名）

### merge 示例

```bash
# 合并两个文件，输出到 merge.json
excel2json merge a.xlsx b.xlsx

# 指定输出文件
excel2json merge a.xlsx b.xlsx c.xlsx -o merged.json

# 格式化输出 JSON 数组
excel2json merge a.xlsx b.xlsx -o out.json -l -p

# 指定 sheet 名
excel2json merge a.xlsx b.xlsx -s "Sheet2" -o out.json

# 关闭来源标记
excel2json merge a.xlsx b.xlsx -o out.json --no-source
```

### 输出格式

默认输出 **NDJSON**（Newline Delimited JSON），每行一个 JSON 对象，适合流式处理和大文件场景：

```
{"姓名":"张三","年龄":"18","城市":"北京"}
{"姓名":"李四","年龄":"20","城市":"上海"}
```

加 `--list` 输出标准 **JSON 数组**：

```json
[
  {"姓名":"张三","年龄":"18","城市":"北京"},
  {"姓名":"李四","年龄":"20","城市":"上海"}
]
```

## 支持格式

| 格式 | 说明 |
|------|------|
| `.xlsx` | Excel 2007+ |
| `.xls` | Excel 97-2003，自动处理 GBK 编码 |
