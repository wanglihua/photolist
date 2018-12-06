package main

import (
	"os"
	"path/filepath"
	"sort"
)

func getPhotoFileList(photoDir string) []string {

	file_list, _ := filepath.Glob(photoDir + string(os.PathSeparator) + "*.*")

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
