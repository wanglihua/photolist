package main

import "os"

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
