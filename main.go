package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/bytedance/sonic"
	xlsReader "github.com/shakinm/xlsReader/xls"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

var (
	sheetName string
	pretty    bool
	noHeader  bool
	list      bool

	// merge flags
	mergeSheet    string
	mergePretty   bool
	mergeNoHeader bool
	mergeList     bool
	mergeOutput   string
	noSource      bool
)

var rootCmd = &cobra.Command{
	Use:   "excel2json <input> <output.json>",
	Short: "将 Excel 文件（.xlsx / .xls）转换为 JSON",
	Args:  cobra.ExactArgs(2),
	RunE:  run,
}

var mergeCmd = &cobra.Command{
	Use:   "merge [flags] <file1.xlsx> <file2.xlsx> ...",
	Short: "合并多个 Excel 文件为一个 JSON 文件（默认输出到 merge.json）",
	Args:  cobra.MinimumNArgs(1),
	RunE:  runMerge,
}

func init() {
	rootCmd.Flags().StringVarP(&sheetName, "sheet", "s", "", "指定读取的工作表名称（默认读取第一个表）")
	rootCmd.Flags().BoolVarP(&pretty, "pretty", "p", false, "输出格式化的 JSON（仅对 --list 有效）")
	rootCmd.Flags().BoolVarP(&noHeader, "no-header", "n", false, "不将第一行作为字段名，改用列索引（col0, col1, ...）")
	rootCmd.Flags().BoolVarP(&list, "list", "l", false, "输出 JSON 数组（默认输出 NDJSON，每行一个对象）")

	mergeCmd.Flags().StringVarP(&mergeSheet, "sheet", "s", "", "指定读取的工作表名称（默认读取第一个表）")
	mergeCmd.Flags().BoolVarP(&mergePretty, "pretty", "p", false, "输出格式化的 JSON（仅对 --list 有效）")
	mergeCmd.Flags().BoolVarP(&mergeNoHeader, "no-header", "n", false, "不将第一行作为字段名，改用列索引（col0, col1, ...）")
	mergeCmd.Flags().BoolVarP(&mergeList, "list", "l", false, "输出 JSON 数组（默认输出 NDJSON，每行一个对象）")
	mergeCmd.Flags().StringVarP(&mergeOutput, "output", "o", "", "输出文件路径")
	mergeCmd.Flags().BoolVar(&noSource, "no-source", false, "关闭来源标记（默认添加 _source 字段）")

	rootCmd.AddCommand(mergeCmd)
}

func run(cmd *cobra.Command, args []string) error {
	inputFile := args[0]
	outputFile := args[1]

	ext := strings.ToLower(filepath.Ext(inputFile))

	var rows [][]string
	var err error

	switch ext {
	case ".xls":
		rows, err = readXLS(inputFile, sheetName)
	case ".xlsx", ".xlsm", "":
		rows, err = readXLSX(inputFile, sheetName)
	default:
		return fmt.Errorf("不支持的文件格式: %s（支持 .xlsx / .xls）", ext)
	}

	if err != nil {
		return err
	}

	result, fieldOrder := rowsToJSON(rows, noHeader)
	if list {
		return writeJSON(outputFile, result, pretty, fieldOrder)
	}
	return writeNDJSON(outputFile, result, fieldOrder)
}

// readXLSX 使用 excelize 读取 .xlsx 文件
func readXLSX(path, sheet string) ([][]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %w", err)
	}
	defer f.Close()

	if sheet == "" {
		sheets := f.GetSheetList()
		if len(sheets) == 0 {
			return nil, fmt.Errorf("文件中没有工作表")
		}
		sheet = sheets[0]
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("读取工作表 %q 失败: %w", sheet, err)
	}
	return rows, nil
}

// readXLS 使用 shakinm/xlsReader 读取旧版 .xls 文件
func readXLS(path, sheet string) ([][]string, error) {
	wb, err := xlsReader.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %w", err)
	}

	// 定位目标 sheet
	var target *xlsReader.Sheet
	for i := 0; i < wb.GetNumberSheets(); i++ {
		s, err := wb.GetSheet(i)
		if err != nil {
			continue
		}
		if sheet == "" || s.GetName() == sheet {
			target = s
			break
		}
	}

	if target == nil {
		if sheet != "" {
			return nil, fmt.Errorf("找不到工作表 %q", sheet)
		}
		return nil, fmt.Errorf("文件中没有工作表")
	}

	var rows [][]string
	for _, row := range target.GetRows() {
		cols := row.GetCols()
		cells := make([]string, len(cols))
		for i, cell := range cols {
			cells[i] = strings.TrimSpace(cell.GetString())
		}
		rows = append(rows, cells)
	}
	return rows, nil
}

// rowsToJSON 将二维字符串表格转换为 JSON 对象列表，返回数据和字段顺序
func rowsToJSON(rows [][]string, noHeader bool) ([]map[string]any, []string) {
	if len(rows) == 0 {
		return []map[string]any{}, nil
	}

	var result []map[string]any
	var fieldOrder []string

	if noHeader {
		// 找出最大列数
		maxCols := 0
		for _, row := range rows {
			if len(row) > maxCols {
				maxCols = len(row)
			}
		}
		// 生成字段顺序
		fieldOrder = make([]string, maxCols)
		for i := 0; i < maxCols; i++ {
			fieldOrder[i] = fmt.Sprintf("col%d", i)
		}
		// 转换数据
		for _, row := range rows {
			record := make(map[string]any, len(fieldOrder))
			for i, key := range fieldOrder {
				if i < len(row) {
					record[key] = row[i]
				} else {
					record[key] = ""
				}
			}
			result = append(result, record)
		}
	} else {
		headers := rows[0]
		fieldOrder = headers
		for _, row := range rows[1:] {
			record := make(map[string]any, len(headers))
			for i, h := range headers {
				if i < len(row) {
					record[h] = row[i]
				} else {
					record[h] = ""
				}
			}
			result = append(result, record)
		}
	}

	if result == nil {
		return []map[string]any{}, fieldOrder
	}
	return result, fieldOrder
}

func runMerge(cmd *cobra.Command, args []string) error {
	// 使用 -o 参数，默认为 merge.json
	outputFile := mergeOutput
	if outputFile == "" {
		outputFile = "merge.json"
	}
	inputFiles := args

	if len(inputFiles) == 0 {
		return fmt.Errorf("至少需要一个输入文件")
	}

	// 收集所有文件的表头和数据
	var allHeaders []string
	var allRecords []map[string]any

	for _, inputFile := range inputFiles {
		ext := strings.ToLower(filepath.Ext(inputFile))

		var rows [][]string
		var err error

		switch ext {
		case ".xls":
			rows, err = readXLS(inputFile, mergeSheet)
		case ".xlsx", ".xlsm", "":
			rows, err = readXLSX(inputFile, mergeSheet)
		default:
			return fmt.Errorf("不支持的文件格式: %s（支持 .xlsx / .xls）", ext)
		}

		if err != nil {
			return fmt.Errorf("读取文件 %s 失败: %w", inputFile, err)
		}

		if len(rows) == 0 {
			continue
		}

		// 获取当前文件的表头
		var headers []string
		if mergeNoHeader {
			// 无表头模式，找出最大列数
			maxCols := 0
			for _, row := range rows {
				if len(row) > maxCols {
					maxCols = len(row)
				}
			}
			headers = make([]string, maxCols)
			for i := 0; i < maxCols; i++ {
				headers[i] = fmt.Sprintf("col%d", i)
			}
		} else {
			headers = rows[0]
		}

		// 合并表头（取并集，保持顺序）
		allHeaders = mergeHeaders(allHeaders, headers)

		// 转换数据
		var records []map[string]any
		if mergeNoHeader {
			for _, row := range rows {
				record := make(map[string]any, len(headers))
				for i, h := range headers {
					if i < len(row) {
						record[h] = row[i]
					} else {
						record[h] = ""
					}
				}
				records = append(records, record)
			}
		} else {
			for _, row := range rows[1:] {
				record := make(map[string]any, len(headers))
				for i, h := range headers {
					if i < len(row) {
						record[h] = row[i]
					} else {
						record[h] = ""
					}
				}
				records = append(records, record)
			}
		}

		// 添加来源标记
		if !noSource {
			source := filepath.Base(inputFile)
			source = strings.TrimSuffix(source, filepath.Ext(inputFile))
			for i := range records {
				records[i]["_source"] = source
			}
		}

		allRecords = append(allRecords, records...)
	}

	// 补全所有记录的字段（确保所有记录都有完整的表头）
	for _, record := range allRecords {
		for _, h := range allHeaders {
			if _, exists := record[h]; !exists {
				record[h] = ""
			}
		}
	}

	// 构建字段顺序：Excel 表头 + _source（如果有）
	fieldOrder := allHeaders
	if !noSource {
		fieldOrder = append([]string{}, allHeaders...)
		fieldOrder = append(fieldOrder, "_source")
	}

	// 输出
	if mergeList {
		return writeJSON(outputFile, allRecords, mergePretty, fieldOrder)
	}
	return writeNDJSON(outputFile, allRecords, fieldOrder)
}

// mergeHeaders 合并两个表头，取并集，保持顺序
func mergeHeaders(existing, new []string) []string {
	seen := make(map[string]bool)
	for _, h := range existing {
		seen[h] = true
	}
	for _, h := range new {
		if !seen[h] {
			existing = append(existing, h)
			seen[h] = true
		}
	}
	return existing
}

func writeNDJSON(path string, data []map[string]any, fieldOrder []string) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("写入文件失败: %w", err)
	}
	defer f.Close()

	for _, record := range data {
		b, err := marshalOrdered(record, fieldOrder)
		if err != nil {
			return fmt.Errorf("序列化 JSON 失败: %w", err)
		}
		f.Write(b)
		f.WriteString("\n")
	}

	fmt.Printf("已写入 %s\n", path)
	return nil
}

func writeJSON(path string, data any, pretty bool, fieldOrder []string) error {
	var b []byte
	var err error

	if pretty {
		if records, ok := data.([]map[string]any); ok && len(fieldOrder) > 0 {
			// 按顺序输出数组
			b, err = marshalOrderedArray(records, fieldOrder, true)
		} else {
			b, err = sonic.MarshalIndent(data, "", "  ")
		}
	} else {
		if records, ok := data.([]map[string]any); ok && len(fieldOrder) > 0 {
			// 按顺序输出数组
			b, err = marshalOrderedArray(records, fieldOrder, false)
		} else {
			b, err = sonic.Marshal(data)
		}
	}

	if err != nil {
		return fmt.Errorf("序列化 JSON 失败: %w", err)
	}

	if err := os.WriteFile(path, b, 0o644); err != nil {
		return fmt.Errorf("写入文件失败: %w", err)
	}

	fmt.Printf("已写入 %s\n", path)
	return nil
}

// marshalOrdered 按指定顺序序列化单个对象
func marshalOrdered(record map[string]any, fieldOrder []string) ([]byte, error) {
	var sb strings.Builder
	sb.WriteString("{")

	first := true
	for _, key := range fieldOrder {
		if val, exists := record[key]; exists {
			if !first {
				sb.WriteString(",")
			}
			first = false

			keyBytes, _ := sonic.Marshal(key)
			valBytes, _ := sonic.Marshal(val)
			sb.Write(keyBytes)
			sb.WriteString(":")
			sb.Write(valBytes)
		}
	}

	sb.WriteString("}")
	return []byte(sb.String()), nil
}

// marshalOrderedArray 按指定顺序序列化对象数组
func marshalOrderedArray(records []map[string]any, fieldOrder []string, pretty bool) ([]byte, error) {
	var sb strings.Builder
	sb.WriteString("[")

	for i, record := range records {
		if i > 0 {
			sb.WriteString(",")
		}
		if pretty {
			sb.WriteString("\n  ")
		}

		sb.WriteString("{")

		first := true
		for _, key := range fieldOrder {
			if val, exists := record[key]; exists {
				if !first {
					sb.WriteString(",")
					if pretty {
						sb.WriteString(" ")
					}
				}
				first = false

				keyBytes, _ := sonic.Marshal(key)
				valBytes, _ := sonic.Marshal(val)
				sb.Write(keyBytes)
				sb.WriteString(":")
				sb.Write(valBytes)
			}
		}

		sb.WriteString("}")
	}

	if pretty {
		sb.WriteString("\n")
	}
	sb.WriteString("]")
	return []byte(sb.String()), nil
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
