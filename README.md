# excel2json

将 Excel 文件（`.xlsx` / `.xls`）转换为 JSON 的命令行工具。

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
