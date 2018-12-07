package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/go-ini/ini"
	"log"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
)

var INI_FILE *ini.File
var WORK_DIR = ""

func main() {
	exeFileName, err := os.Executable()
	if err != nil {
		fmt.Println(err)
	}

	WORK_DIR, err = filepath.Abs(filepath.Dir(exeFileName))
	if err != nil {
		log.Println(err)
	}

	var iniFileName = "photolist.ini"
	var iniFileFullName = WORK_DIR + string(os.PathSeparator) + iniFileName

	if !pathExists(iniFileFullName) {
		var goPath = os.Getenv("GOPATH")
		iniFileFullName = goPath + string(os.PathSeparator) + "src" + string(os.PathSeparator) + "photolist" + string(os.PathSeparator) + iniFileName
	}

	if !pathExists(iniFileFullName) {
		fmt.Println("photolist.ini file not exist!")
	}

	INI_FILE, err = ini.Load(iniFileFullName)
	if err != nil {
		log.Println(err)
	}

	var photoDir = INI_FILE.Section("").Key("photo_dir").String()
	var excelFile = INI_FILE.Section("").Key("excel_file").String()
	var excelSheetName = INI_FILE.Section("").Key("excel_sheet_name").String()
	var excelPhotoNameCol = INI_FILE.Section("").Key("excel_photo_name_col").String()
	var excelRemarkCol = INI_FILE.Section("").Key("excel_remark_col").String()
	var remarkText = INI_FILE.Section("").Key("remark_text").String()

	remarkInt, errRemarkInt := strconv.ParseInt(remarkText, 10, 64)

	// get photo file list
	var photoFileList = getPhotoFileList(photoDir)

	var photoFileMap = make(map[string]bool)
	for _, photoFile := range photoFileList {
		var photoFileKey = strings.TrimSuffix(photoFile, path.Ext(photoFile))
		photoFileKey = strings.TrimSpace(photoFileKey)
		photoFileKey = strings.ToUpper(photoFileKey)

		photoFileMap[photoFileKey] = false
	}

	xlsx, err := excelize.OpenFile(excelFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := xlsx.GetRows(excelSheetName)
	for index, _ := range rows {
		if index == 0 { // 略过标题行
			continue
		}

		var photoNameCellName = fmt.Sprintf("%s%d", excelPhotoNameCol, index+1)
		var remarkCellName = fmt.Sprintf("%s%d", excelRemarkCol, index+1)

		var photoNameCellValue = xlsx.GetCellValue(excelSheetName, photoNameCellName)
		photoNameCellValue = strings.TrimSpace(photoNameCellValue)
		photoNameCellValue = strings.ToUpper(photoNameCellValue)

		_, photoFileExist := photoFileMap[photoNameCellValue]
		//设置标记
		if photoFileExist == true {
			photoFileMap[photoNameCellValue] = true
			if errRemarkInt != nil { // 是文本
				xlsx.SetCellStr(excelSheetName, remarkCellName, remarkText)
			} else { // 是数字
				xlsx.SetCellInt(excelSheetName, remarkCellName, int(remarkInt))
			}
		}

	}

	var resultSheetName = createResultSheet(xlsx)
	var resultFoundCellNum = 2;
	xlsx.SetCellStr(resultSheetName, "A1", "已找到")
	var resultNotFoundCellNum = 2;
	xlsx.SetCellStr(resultSheetName, "B1", "未找到")

	for _, photoFile := range photoFileList {
		var photoFileKey = strings.TrimSuffix(photoFile, path.Ext(photoFile))
		photoFileKey = strings.TrimSpace(photoFileKey)
		photoFileKey = strings.ToUpper(photoFileKey)

		var photoFileExist = photoFileMap[photoFileKey]
		// 输出结果
		if photoFileExist {
			xlsx.SetCellStr(resultSheetName, fmt.Sprintf("A%d", resultFoundCellNum), photoFile)
			resultFoundCellNum = resultFoundCellNum + 1
		} else {
			xlsx.SetCellStr(resultSheetName, fmt.Sprintf("B%d", resultNotFoundCellNum), photoFile)
			resultNotFoundCellNum = resultNotFoundCellNum + 1
		}
	}

	err = xlsx.Save()
	if err != nil {
		fmt.Println(err)
	}
}

func pathExists(path string) bool {
	_, err := os.Stat(path)
	if err == nil {
		// return true, nil
		return true
	}

	if os.IsNotExist(err) {
		// return false, nil
		return false
	}

	// return false, err
	return false
}

func getPhotoFileList(photoDir string) []string {

	fileList, _ := filepath.Glob(photoDir + string(os.PathSeparator) + "*.*")

	for i := 0; i < len(fileList); i++ {
		fileList[i] = filepath.Base(fileList[i])
	}

	var fileNameList = FileNameList(fileList)
	// sort.Sort(sort.Reverse(fileNameList))
	sort.Sort(fileNameList)

	fileList = []string(fileNameList)

	return fileList
}

type FileNameList []string

func (fileNameList FileNameList) Len() int {
	return len(fileNameList)
}

func (fileNameList FileNameList) Less(i, j int) bool {
	return fileNameList[i] < fileNameList[j]
}

func (fileNameList FileNameList) Swap(i, j int) {
	fileNameList[i], fileNameList[j] = fileNameList[j], fileNameList[i]
}

func createResultSheet(xlsx *excelize.File) string {
	var sheetNameMap = make(map[string]bool)
	for _, name := range xlsx.GetSheetMap() {
		sheetNameMap[name] = true
	}

	var sheetNameBase = "查找结果"

	_, found := sheetNameMap[sheetNameBase]
	if !found {
		xlsx.NewSheet(sheetNameBase)
		return sheetNameBase
	}

	for i := 2; i < 100; i ++ {
		var resultSheetName = fmt.Sprintf("%s%d", sheetNameBase, i)
		_, found := sheetNameMap[resultSheetName]
		if !found {
			xlsx.NewSheet(resultSheetName)
			return resultSheetName
		}
	}

	return sheetNameBase
}
