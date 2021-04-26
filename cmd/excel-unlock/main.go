package main

import (
	"fmt"
	"os"

	"github.com/Shamus03/excel_unlocker"
)

func main() {
	if err := mainErr(); err != nil {
		fmt.Fprintf(os.Stderr, err.Error())
		os.Exit(1)
	}
}

func mainErr() error {
	if len(os.Args) <= 1 {
		return fmt.Errorf("missing file name")
	}

	return excel_unlocker.UnlockFile(os.Args[1])
}
