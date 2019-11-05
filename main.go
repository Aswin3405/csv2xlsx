package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/jiazhoulvke/goutil"
	"github.com/spf13/pflag"
)

const (
	ColumnTypeInt    = "i"
	ColumnTypeFloat  = "f"
	ColumnTypeString = "s"
)

var (
	columnsType arrayFlags
	ColumnsType map[int]string
	Header      string = "1"
	Output      string = "./"
)

type arrayFlags []string

func (p *arrayFlags) Set(value string) error {
	*p = append(*p, value)
	return nil
}

func (p *arrayFlags) String() string {
	return ""
}

func (p *arrayFlags) Type() string {
	return "[]string"
}

func init() {
	ColumnsType = make(map[int]string)

	pflag.StringVarP(&Output, "output", "o", Output, "output dir")
	pflag.StringVar(&Header, "header", Header, `set header
0:no header 1:has header
`)
	pflag.VarP(&columnsType, "column", "c", `set column type
i:int, f:float, s:string
example: -cA:i -c2:s --column=C:f`)
}

func main() {
	pflag.Parse()
	if !goutil.InStringSlice(Header, []string{"0", "1"}) {
		fmt.Println("header error: must be 0 or 1")
		os.Exit(1)
	}

	reNumber := regexp.MustCompile(`\d+`)
	for _, ct := range columnsType {
		l := strings.Split(ct, ":")
		if len(l) != 2 {
			fmt.Println("set column type:", ct)
			os.Exit(1)
		}
		var column int
		var err error
		if reNumber.Match([]byte(l[0])) {
			column, err = strconv.Atoi(l[0])
		} else {
			column, err = excelize.ColumnNameToNumber(l[0])
		}
		if err != nil {
			fmt.Println("set column type:", err)
			os.Exit(1)
		}
		ctlist := []string{ColumnTypeInt, ColumnTypeFloat, ColumnTypeString}
		if !goutil.InStringSlice(l[1], ctlist) {
			fmt.Println("column type is not supported:", l[1], ", must in:", ctlist)
			os.Exit(1)
		}
		ColumnsType[column] = l[1]
	}
	if len(pflag.Args()) == 0 {
		fmt.Println("input file is required")
		os.Exit(1)
	}

	for _, filename := range pflag.Args() {
		err := ConvertCSV2XLSX(filename)
		if err != nil {
			fmt.Printf("convert file [%s] error: %v\n", filename, err)
			os.Exit(1)
		}
	}
}

func ConvertCSV2XLSX(filename string) error {
	f, err := os.Open(filename)
	if err != nil {
		return err
	}
	defer f.Close()
	r := csv.NewReader(f)
	var isHeader bool
	if Header == "1" {
		isHeader = true
	}
	excel := excelize.NewFile()
	sheet := "Sheet1"
	i := 1
	for {
		line, err := r.Read()
		if err != nil {
			if err != io.EOF {
				return err
			}
			break
		}
		for c, cellValue := range line {
			cellNumber := c + 1
			cellName, err := excelize.CoordinatesToCellName(cellNumber, i)
			if err != nil {
				return err
			}
			cellType := ColumnsType[cellNumber]
			if isHeader {
				cellType = ""
			}
			switch cellType {
			case ColumnTypeInt:
				n, err := strconv.Atoi(cellValue)
				if err != nil {
					return fmt.Errorf("can't convert cell[%s]: %v", cellName, err)
				}
				if err := excel.SetCellInt(sheet, cellName, n); err != nil {
					return err
				}
			case ColumnTypeFloat:
				n, err := strconv.ParseFloat(cellValue, 64)
				if err != nil {
					return fmt.Errorf("can't convert cell[%s]: %v", cellName, err)
				}
				if err := excel.SetCellFloat(sheet, cellName, n, -1, 64); err != nil {
					return err
				}
			case ColumnTypeString:
				if err := excel.SetCellStr(sheet, cellName, cellValue); err != nil {
					return err
				}
			default:
				if err := excel.SetCellDefault(sheet, cellName, cellValue); err != nil {
					return err
				}
			}
		}
		if isHeader {
			isHeader = false
		}
		i++
	}
	bname := filepath.Base(filename)
	xlsxFilename := strings.TrimRight(bname, filepath.Ext(bname)) + ".xlsx"
	savePath := filepath.Join(Output, xlsxFilename)
	return excel.SaveAs(savePath)
}
