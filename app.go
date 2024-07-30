package main

import (
	"bufio"
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

// App struct
type App struct {
	ctx context.Context
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

func (a *App) SelectFile(title, displayName, pattern string) (string, error) {
	dialog := runtime.OpenDialogOptions{
		Title: title,
		Filters: []runtime.FileFilter{
			{DisplayName: displayName, Pattern: pattern},
		},
	}
	filePath, err := runtime.OpenFileDialog(a.ctx, dialog)
	if err != nil {
		return "", err
	}
	return filePath, nil
}

func (a *App) ReplaceTextInSheet(excelFilePath, ruleFilePath string) string {

	f, err := openExcelFile(excelFilePath)
	if err != nil {

		return fmt.Sprintf("Error opening Excel file: %v\n", err)
	}

	basePath := filepath.Dir(f.Path)

	ruleFile, err := openFile(ruleFilePath)
	if err != nil {

		return fmt.Sprintf("Error opening Rule file: %v\n", err)
	}

	rules, err := getReplaceRule(ruleFile)
	if err != nil {

		return fmt.Sprintf("Error scanning Rule file: %v\n", err)
	}

	modifyFile := filepath.Join(basePath, "modified.xlsx")

	defer saveAndCloseExcel(f, modifyFile)

	for _, sheetName := range f.GetSheetMap() {
		for _, rule := range rules {
			oldText, newText := rule[0], rule[1]

			if err := replaceTextInSheet(f, sheetName, oldText, newText); err != nil {

				return fmt.Sprintf("Error replacing text in sheet '%s': %v\n", sheetName, err)
			}
		}
	}

	return fmt.Sprintf("文本替换完成，已保存到 %s", modifyFile)
}

// 打开并验证Excel文件
func openExcelFile(fileName string) (*excelize.File, error) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return nil, err
	}
	return f, nil
}

func openFile(fileName string) (*os.File, error) {
	f, err := os.Open(fileName)
	if err != nil {
		return nil, err
	}
	return f, nil
}

func getReplaceRule(f *os.File) ([][]string, error) {
	scanner := bufio.NewScanner(f)
	if scanner.Err() != nil {
		return nil, scanner.Err()
	}
	var batchData [][]string
	for scanner.Scan() {
		batchData = append(batchData, strings.Split(scanner.Text(), " "))
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Print("关闭文件错误")
		}
	}()

	return batchData, nil
}

// 替换工作表中的文本
func replaceTextInSheet(f *excelize.File, sheetName, oldText, newText string) error {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}
	for rowIndex, row := range rows {
		for columnIndex, cell := range row {
			if cell == oldText {
				columnName, _ := excelize.ColumnNumberToName(columnIndex + 1)
				cellCoordinates := columnName + strconv.Itoa(rowIndex+1)
				f.SetCellValue(sheetName, cellCoordinates, newText)
			}
		}
	}
	return nil
}

// 保存并关闭Excel文件
func saveAndCloseExcel(f *excelize.File, fileName string) error {
	if err := f.SaveAs(fileName); err != nil {
		return err
	}
	if err := f.Close(); err != nil {
		return err
	}
	return nil
}
