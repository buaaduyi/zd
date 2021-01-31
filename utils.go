package main

import (
	"fmt"
	"io/ioutil"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// 学生种类数量
const studentKindNum = 6

// sheet 名称
const sheetName string = "Sheet1"

// StudentInfo 学生信息
type StudentInfo struct {
	name string
	id   string
}

var studentList []StudentInfo

// XlsxInfo 数据表信息
type XlsxInfo struct {
	file  *excelize.File
	index int
	count int
}

var xlsxList []*XlsxInfo

var xlsxName = []string{"博士生未打卡名单", "研究生未打卡名单",
	"2017级本科生未打卡名单", "2018级本科生未打卡名单",
	"2019级本科生未打卡名单", "2020级本科生未打卡名单"}

func readExcel(filePath string) {

	xlsx, err := excelize.OpenFile(filePath)
	errCheck(err)

	rows := xlsx.GetRows(xlsx.GetSheetName(1))
	if len(rows) > 0 {
		rows = rows[1:]
	}

	for _, row := range rows {
		if len(row) >= 3 {
			studentInfo := StudentInfo{
				name: row[1],
				id:   row[2],
			}
			studentList = append(studentList, studentInfo)
		}
	}
}

func filter() {

	xlsxList = xlsxList[0:0]

	if len(xlsxList) != 0 {
		return
	}

	for i := 0; i < studentKindNum; i++ {
		file := excelize.NewFile()
		index := file.NewSheet(sheetName)

		xlsxInfo := XlsxInfo{
			file:  file,
			index: index,
			count: 1,
		}

		xlsxList = append(xlsxList, &xlsxInfo)
	}

	boshiData = boshiData[0:0]
	shuoshiData = shuoshiData[0:0]
	benke17Data = benke17Data[0:0]
	benke18Data = benke18Data[0:0]
	benke19Data = benke19Data[0:0]
	benke20Data = benke20Data[0:0]

	if len(boshiData) != 0 ||
		len(shuoshiData) != 0 ||
		len(benke17Data) != 0 ||
		len(benke18Data) != 0 ||
		len(benke19Data) != 0 ||
		len(benke20Data) != 0 {
		return
	}

	for _, studentInfo := range studentList {
		id := studentInfo.id
		if len(id) >= 3 {
			if id[2] == '1' {
				xlsxInfo := xlsxList[0]
				countStr := strconv.Itoa(xlsxInfo.count)
				xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
				xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
				temp := []string{countStr, studentInfo.name, studentInfo.id}
				boshiData = append(boshiData, temp)
				xlsxInfo.count++
			} else if id[2] == '2' {
				xlsxInfo := xlsxList[1]
				countStr := strconv.Itoa(xlsxInfo.count)
				xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
				xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
				temp := []string{countStr, studentInfo.name, studentInfo.id}
				shuoshiData = append(shuoshiData, temp)
				xlsxInfo.count++
			} else if id[2] == '3' {
				if id[:2] == "17" {
					xlsxInfo := xlsxList[2]
					countStr := strconv.Itoa(xlsxInfo.count)
					xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
					xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
					temp := []string{countStr, studentInfo.name, studentInfo.id}
					benke17Data = append(benke17Data, temp)
					xlsxInfo.count++
				} else if id[:2] == "18" {
					xlsxInfo := xlsxList[3]
					countStr := strconv.Itoa(xlsxInfo.count)
					xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
					xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
					temp := []string{countStr, studentInfo.name, studentInfo.id}
					benke18Data = append(benke18Data, temp)
					xlsxInfo.count++
				} else if id[:2] == "19" {
					xlsxInfo := xlsxList[4]
					countStr := strconv.Itoa(xlsxInfo.count)
					xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
					xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
					temp := []string{countStr, studentInfo.name, studentInfo.id}
					benke19Data = append(benke19Data, temp)
					xlsxInfo.count++
				} else if id[:2] == "20" {
					xlsxInfo := xlsxList[5]
					countStr := strconv.Itoa(xlsxInfo.count)
					xlsxInfo.file.SetCellValue(sheetName, "A"+countStr, studentInfo.name)
					xlsxInfo.file.SetCellValue(sheetName, "B"+countStr, studentInfo.id)
					temp := []string{countStr, studentInfo.name, studentInfo.id}
					benke20Data = append(benke20Data, temp)
					xlsxInfo.count++
				}
			}
		}
	}

	if len(xlsxName) != studentKindNum || len(xlsxList) != studentKindNum {
		return
	}

	now := time.Now()
	timeStr := now.Format("20060102") + "日"

	for i := 0; i < len(xlsxName); i++ {
		xlsxName[i] = timeStr + xlsxName[i] + ".xlsx"
	}

	for i := 0; i < studentKindNum; i++ {
		xlsxInfo := xlsxList[i]
		xlsxInfo.file.SetActiveSheet(xlsxInfo.index)
		xlsxInfo.file.SaveAs(OutputPath + xlsxName[i])
	}

}

func errCheck(err error) {
	if err != nil {
		fmt.Println("ERROR: " + err.Error())
	}
}

func getFileList(dirPath string) []string {
	var fileNames []string
	fileList, err := ioutil.ReadDir(dirPath)
	errCheck(err)
	for _, file := range fileList {
		fileNames = append(fileNames, file.Name())
	}
	return fileNames
}
