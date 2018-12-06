package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/go-ini/ini"
	"log"
	"os"
	"path/filepath"
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
			xlsx.SetCellStr(excelSheetName, remarkCellName, remarkText)
		}
	}
}

func isNameInPhotoFileList(name string, photoFileList []string) bool {

	for _, photoFile := range photoFileList {
		if strings.ToUpper(strings.TrimSpace(name)) == strings.ToUpper(strings.TrimSpace(photoFile)) {
			return true
		}
	}

	return false
}
