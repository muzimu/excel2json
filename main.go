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
)

var rootCmd = &cobra.Command{
	Use:   "excel2json <input> <output.json>",
	Short: "将 Excel 文件（.xlsx / .xls）转换为 JSON",
	Args:  cobra.ExactArgs(2),
	RunE:  run,
}

func init() {
	rootCmd.Flags().StringVarP(&sheetName, "sheet", "s", "", "指定读取的工作表名称（默认读取第一个表）")
	rootCmd.Flags().BoolVarP(&pretty, "pretty", "p", false, "输出格式化的 JSON（仅对 --list 有效）")
	rootCmd.Flags().BoolVarP(&noHeader, "no-header", "n", false, "不将第一行作为字段名，改用列索引（col0, col1, ...）")
	rootCmd.Flags().BoolVarP(&list, "list", "l", false, "输出 JSON 数组（默认输出 NDJSON，每行一个对象）")
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

	result := rowsToJSON(rows, noHeader)
	if list {
		return writeJSON(outputFile, result, pretty)
	}
	return writeNDJSON(outputFile, result)
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

// rowsToJSON 将二维字符串表格转换为 JSON 对象列表
func rowsToJSON(rows [][]string, noHeader bool) []map[string]any {
	if len(rows) == 0 {
		return []map[string]any{}
	}

	var result []map[string]any

	if noHeader {
		for _, row := range rows {
			record := make(map[string]any, len(row))
			for i, cell := range row {
				record[fmt.Sprintf("col%d", i)] = cell
			}
			result = append(result, record)
		}
	} else {
		headers := rows[0]
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
		return []map[string]any{}
	}
	return result
}

func writeNDJSON(path string, data []map[string]any) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("写入文件失败: %w", err)
	}
	defer f.Close()

	for _, record := range data {
		b, err := sonic.Marshal(record)
		if err != nil {
			return fmt.Errorf("序列化 JSON 失败: %w", err)
		}
		f.Write(b)
		f.WriteString("\n")
	}

	fmt.Printf("已写入 %s\n", path)
	return nil
}

func writeJSON(path string, data any, pretty bool) error {
	var (
		b   []byte
		err error
	)
	if pretty {
		b, err = sonic.MarshalIndent(data, "", "  ")
	} else {
		b, err = sonic.Marshal(data)
	}
	if err != nil {
		return fmt.Errorf("序列化 JSON 失败: %w", err)
	}

	if err := os.WriteFile(path, b, 0644); err != nil {
		return fmt.Errorf("写入文件失败: %w", err)
	}

	fmt.Printf("已写入 %s\n", path)
	return nil
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}
