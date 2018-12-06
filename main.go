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

	xlsx, err := excelize.OpenFile(excelFile)
	if err != nil {
		fmt.Println(err)
		return
	}

	// Get all the rows in the Sheet1.
	rows := xlsx.GetRows(excelSheetName)
	for index, _ := range rows {
		var photoNameCellName = fmt.Sprintf("%s%d", excelPhotoNameCol, index+1)
		var remarkCellName = fmt.Sprintf("%s%d", excelRemarkCol, index+1)

		var photoNameCellValue = xlsx.GetCellValue(excelSheetName, photoNameCellName)

		if isNameInPhotoFileList(photoNameCellValue, photoFileList) {
			if errRemarkInt != nil { // 是文本
				xlsx.SetCellStr(excelSheetName, remarkCellName, remarkText)
			} else { // 是数字
				xlsx.SetCellInt(excelSheetName, remarkCellName, int(remarkInt))
			}
		}
	}

	err = xlsx.Save()
	if err != nil {
		fmt.Println(err)
	}
}

func isNameInPhotoFileList(name string, photoFileList []string) bool {

	for _, photoFile := range photoFileList {
		var fileNameWithoutExt = strings.TrimSuffix(photoFile, path.Ext(photoFile))
		if strings.ToUpper(strings.TrimSpace(name)) == strings.ToUpper(strings.TrimSpace(fileNameWithoutExt)) {
			return true
		}
	}

	return false
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

	file_list, _ := filepath.Glob(photoDir + string(os.PathSeparator) + "*.*")

	for i := 0; i < len(file_list); i++ {
		file_list[i] = filepath.Base(file_list[i])
	}

	var fileNameList = FileNameList(file_list)
	// sort.Sort(sort.Reverse(fileNameList))
	sort.Sort(fileNameList)

	file_list = []string(fileNameList)

	return file_list
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
