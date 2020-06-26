package xlsxconv

import (
	"fmt"
	"mime/multipart"
	"os"
	"strings"

	"github.com/ecarter202/csv2xlsx"
	"github.com/ecarter202/randstr"
	"github.com/plandem/xlsx"
)

func Open(f multipart.File, h *multipart.FileHeader) (xl *xlsx.Spreadsheet, err error) {
	path := fmt.Sprintf("%s/%s_%s.xlsx", os.TempDir(), h.Filename, randstr.Generate(8))

	ext := strings.Split(h.Filename, ".")[len(strings.Split(h.Filename, "."))-1]
	if ext == "csv" {
		if xl, err = csv2xlsx.Convert(f, "Sheet 1"); err != nil {
			return nil, err
		}
	} else {
		if xl, err = xlsx.Open(f); err != nil {
			return nil, fmt.Errorf("error opening xlsx: %v", err)
		}
	}
	if err = xl.SaveAs(path); err != nil {
		return nil, fmt.Errorf("unable to save local xlsx: %v", err)
	}

	return xlsx.Open(path)
}
