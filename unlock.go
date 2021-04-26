package excel_unlocker

import (
	"archive/zip"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

func UnlockFile(fileName string) error {
	zf, err := zip.OpenReader(fileName)
	if err != nil {
		return fmt.Errorf("open Excel file: %v", err)
	}
	defer zf.Close()

	ext := filepath.Ext(fileName)
	of, err := os.Create(strings.TrimSuffix(fileName, ext) + " - Unlocked" + ext)
	if err != nil {
		return fmt.Errorf("create unlocked file: %v", err)
	}
	defer of.Close()

	zw := zip.NewWriter(of)
	defer zw.Close()

	sheetProtectionRegex := regexp.MustCompile(`<sheetProtection [^>]*/>`)

	foundSheet := false
	for _, f := range zf.File {
		fr, err := f.Open()
		if err != nil {
			return fmt.Errorf("open archived file for reading: %v", err)
		}
		defer fr.Close()

		fw, err := zw.Create(f.Name)
		if err != nil {
			return fmt.Errorf("create unlocked archive file: %v", err)
		}

		if filepath.Dir(f.Name) == filepath.Join("xl", "worksheets") {
			foundSheet = true

			bytes, err := ioutil.ReadAll(fr)
			if err != nil {
				return fmt.Errorf("read all text from sheet file: %v", err)
			}

			unlockedBytes := sheetProtectionRegex.ReplaceAll(bytes, nil)

			if _, err := fw.Write(unlockedBytes); err != nil {
				return fmt.Errorf("write unlocked sheet file: %v", err)
			}
		} else {
			if _, err := io.Copy(fw, fr); err != nil {
				return fmt.Errorf("copy archive file: %v", err)
			}
		}
	}

	if !foundSheet {
		return fmt.Errorf("did not find any Excel sheets in the given file")
	}

	return nil
}
