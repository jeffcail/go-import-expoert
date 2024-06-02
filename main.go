package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"net/http"
	"strings"
)

var file *excelize.File

func importData(w http.ResponseWriter, r *http.Request) {
	// 获取请求中的文件名
	formFile, _, err := r.FormFile("filename")
	if err != nil {
		w.Write([]byte("获取文件失败, " + err.Error()))
		return
	}
	// 关闭
	defer formFile.Close()
	//
	reader, err := excelize.OpenReader(formFile)
	if err != nil {
		w.Write([]byte("读取文件失败, " + err.Error()))
		return
	}
	// 关闭
	defer reader.Close()
	rows, err := reader.GetRows("Sheet1")
	if err != nil {
		w.Write([]byte("获取工作表失败, " + err.Error()))
		return
	}
	for i, row := range rows {
		// 每一行数据的列, 都是从0开始的, 一般0行都是表头
		if i == 0 {
			continue
		}
		value1 := row[0] // 第一列
		value2 := row[1] // 第二列
		// 去除空格
		value1 = strings.Trim(strings.TrimSpace(value1), "\n")
		value2 = strings.Trim(strings.TrimSpace(value2), "\n")
		//
		fmt.Fprintf(w, fmt.Sprintf("%s%s\n", value1, value2))
	}
	return
}

func exportData(w http.ResponseWriter, r *http.Request) {
	defer func() {
		file.Close()
	}()

	file = excelize.NewFile()

	// 设置页
	sheetName := "Sheet1"
	// 创建
	sheet, err := file.NewSheet(sheetName)
	if err != nil {
		w.Write([]byte("创建失败, " + err.Error()))
		return
	}
	// 设置单元格格式
	style := setStyle()

	styleID, _ := file.NewStyle(style)
	// 设置表头
	_ = file.SetCellValue(sheetName, "A1", "款")
	_ = file.SetCellStyle(sheetName, "A1", "A1", styleID)
	_ = file.SetCellValue(sheetName, "B1", "尺码")
	_ = file.SetCellStyle(sheetName, "B1", "B1", styleID)
	// 设置值
	for i := 0; i < 10; i++ {
		line := fmt.Sprintf("%d", i+2)

		_ = file.SetCellValue(sheetName, "A"+line, "基础款")
		_ = file.SetCellStyle(sheetName, "A"+line, "A"+line, styleID)

		_ = file.SetCellValue(sheetName, "B"+line, "1:2:3:4:5:6")
		_ = file.SetCellStyle(sheetName, "B"+line, "B"+line, styleID)
	}
	file.SetActiveSheet(sheet)
	buffer, err := file.WriteToBuffer()
	if err != nil {
		w.Write([]byte("导出失败, " + err.Error()))
		return
	}

	// Content-Type:application/octet-stream
	// 导出的文件格式 csv
	// 告知浏览器这是一个二进制字节流，浏览器处理字节流的默认方式就是下载
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "filename.csv"))

	// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
	// 导出的文件格式 xlsx
	//w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "导出文件.xlsx"))
	//w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// application/vnd.ms-excel
	// 导出的文件格式是xls. Microsoft Excel 文件类型的 MIME 类型
	// 需要注意的是，application/vnd.ms-excel 仅适用于旧版本的 Microsoft Excel 文件（.xls）。
	// 对于较新的 .xlsx 文件，应使用 application/vnd.openxmlformats-officedocument.spreadsheetml.sheet 作为 MIME 类型。
	//w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "导出文件.xsl"))
	//w.Header().Set("Content-Type", "application/vnd.ms-excel")

	w.Write(buffer.Bytes())
}

func exportData2(w http.ResponseWriter, r *http.Request) {
	defer func() {
		file.Close()
	}()
	file = excelize.NewFile()

	// 设置页
	sheetName := "Sheet1"
	// 创建
	sheet, err := file.NewSheet(sheetName)
	if err != nil {
		w.Write([]byte("创建失败, " + err.Error()))
		return
	}
	// 设置单元格格式
	style := setStyle()
	styleID, _ := file.NewStyle(style)
	// 设置表头
	_ = file.MergeCell(sheetName, "A1", "A2") // 合并单元格
	_ = file.SetCellValue(sheetName, "A1", "款")
	_ = file.SetCellStyle(sheetName, "A1", "A2", styleID)

	_ = file.MergeCell(sheetName, "B1", "G1") // 合并单元格
	_ = file.SetCellValue(sheetName, "B1", "尺码")
	_ = file.SetCellStyle(sheetName, "B1", "G1", styleID)
	_ = file.SetCellValue(sheetName, "B2", "XS")
	_ = file.SetCellStyle(sheetName, "B2", "B2", styleID)
	_ = file.SetCellValue(sheetName, "C2", "S")
	_ = file.SetCellStyle(sheetName, "C2", "C2", styleID)
	_ = file.SetCellValue(sheetName, "D2", "M")
	_ = file.SetCellStyle(sheetName, "D2", "D2", styleID)
	_ = file.SetCellValue(sheetName, "E2", "L")
	_ = file.SetCellStyle(sheetName, "E2", "E2", styleID)
	_ = file.SetCellValue(sheetName, "F2", "XL")
	_ = file.SetCellStyle(sheetName, "F2", "F2", styleID)
	_ = file.SetCellValue(sheetName, "G2", "XLL")
	_ = file.SetCellStyle(sheetName, "G2", "G2", styleID)
	// 设置值
	for i := 0; i < 10; i++ {
		lineStr := fmt.Sprintf("%d", i+3)
		//
		_ = file.SetCellValue(sheetName, "A"+lineStr, "基础款")
		_ = file.SetCellStyle(sheetName, "A"+lineStr, "A"+lineStr, styleID)
		//
		split := strings.Split("1:2:3:4:5:6", ":")
		_ = file.SetCellValue(sheetName, "B"+lineStr, split[0])
		_ = file.SetCellStyle(sheetName, "B"+lineStr, "B"+lineStr, styleID)
		_ = file.SetCellValue(sheetName, "C"+lineStr, split[1])
		_ = file.SetCellStyle(sheetName, "C"+lineStr, "C"+lineStr, styleID)
		_ = file.SetCellValue(sheetName, "D"+lineStr, split[2])
		_ = file.SetCellStyle(sheetName, "D"+lineStr, "D"+lineStr, styleID)
		_ = file.SetCellValue(sheetName, "E"+lineStr, split[3])
		_ = file.SetCellStyle(sheetName, "E"+lineStr, "E"+lineStr, styleID)
		_ = file.SetCellValue(sheetName, "F"+lineStr, split[4])
		_ = file.SetCellStyle(sheetName, "F"+lineStr, "F"+lineStr, styleID)
		_ = file.SetCellValue(sheetName, "G"+lineStr, split[5])
		_ = file.SetCellStyle(sheetName, "G"+lineStr, "G"+lineStr, styleID)
	}
	//
	file.SetActiveSheet(sheet)
	//
	buffer, err := file.WriteToBuffer()
	if err != nil {
		w.Write([]byte("导出失败, " + err.Error()))
		return
	}

	// Content-Type:application/octet-stream
	// 导出的文件格式 csv
	// 告知浏览器这是一个二进制字节流，浏览器处理字节流的默认方式就是下载
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "filename.csv"))

	// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
	// 导出的文件格式 xlsx
	//w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "导出文件.xlsx"))
	//w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	// application/vnd.ms-excel
	// 导出的文件格式是xls. Microsoft Excel 文件类型的 MIME 类型
	// 需要注意的是，application/vnd.ms-excel 仅适用于旧版本的 Microsoft Excel 文件（.xls）。
	// 对于较新的 .xlsx 文件，应使用 application/vnd.openxmlformats-officedocument.spreadsheetml.sheet 作为 MIME 类型。
	//w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", "导出文件.xsl"))
	//w.Header().Set("Content-Type", "application/vnd.ms-excel")

	w.Write(buffer.Bytes())
}

func setStyle() *excelize.Style {
	return &excelize.Style{
		Border: nil,
		Fill:   excelize.Fill{},
		Font:   nil,
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          0,
			JustifyLastLine: false,
			ReadingOrder:    0,
			RelativeIndent:  0,
			ShrinkToFit:     false,
			TextRotation:    0,
			Vertical:        "center",
			WrapText:        false,
		},
		Protection:    nil,
		NumFmt:        0,
		DecimalPlaces: nil,
		CustomNumFmt:  nil,
		NegRed:        false,
	}
}

const banner = `
   __ _                                  
  / _ |   ___    __ __    ___     __
\__, |  (_-<    \ V /   (_-<    / _|
|___/   /__/_   _\_/_   /__/_   \__|_

`

func main() {
	// HTTP服务
	http.HandleFunc("/import", importData)
	http.HandleFunc("/export", exportData)
	http.HandleFunc("/export2", exportData2)
	fmt.Println(fmt.Sprintf("%s", banner))
	fmt.Println(fmt.Sprintf("run server port: %s", ":9000"))

	err := http.ListenAndServe(":9000", nil)
	if err != nil {
		panic(err)
	}

}
