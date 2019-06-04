// goexcel
package main

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"runtime"

	//"strings"
	"time"

	"github.com/aswjh/excel"
)

type Customer struct {
	no        int
	cname     string
	firstname string
	lastname  string
	birthday  string
	persionid string
}

//取得當前目錄
func getCurrentDirectory() string {
	dir, err := filepath.Abs(filepath.Dir(os.Args[0]))

	fmt.Println(dir)

	if err != nil {
		fmt.Println(err)
	}
	return dir //strings.Replace(dir, "\\", "/", -1) d
}

func mkdir(xdir string) {

	_, err := os.Stat(xdir)
	if err != nil {
		fmt.Println(xdir, "dir error,maybe is not exist, maybe not")
		if os.IsNotExist(err) {
			fmt.Println(xdir, "dir is not exist")
			err := os.Mkdir(xdir, os.ModePerm)
			if err != nil {
				fmt.Printf("mkdir failed![%v]\n", err)
			}
			return
		}

		fmt.Println("stat file error")
		return
	}

	fmt.Println(xdir, "dir is exist")
}

func main() {
	runtime.GOMAXPROCS(1)

	sSysDir := getCurrentDirectory()
	sOldPath := sSysDir + "\\old"
	sNewPath := sSysDir + "\\new"

	//建立目錄
	//mkdir(sOldPath)
	os.MkdirAll(sOldPath, os.ModePerm)
	os.RemoveAll(sNewPath)
	os.MkdirAll(sNewPath, os.ModePerm)

	option := excel.Option{"Visible": false, "DisplayAlerts": true, "ScreenUpdating": true}
	xl, _ := excel.New(option) //xl, _ := excel.Open("test_excel.xls", option)

	//defer xl.Quit()

	sheet, _ := xl.Sheet(1) //xl.Sheet("sheet1")
	//defer sheet.Release()
	sheet.Cells(1, 1, "hello")
	sheet.PutCell(1, 2, 2006)
	sheet.MustCells(1, 3, 3.14159)

	cell := sheet.Cell(5, 6)
	//defer cell.Release()
	cell.Put("go")
	cell.Put("font", map[string]interface{}{"name": "Arial", "size": 26, "bold": true})
	cell.Put("interior", "colorindex", 6)

	sheet.PutRange("a3:c3", []string{"@1", "@2", "@3"})
	rg := sheet.Range("d3:f3")
	//defer rg.Release()
	rg.Put([]string{"~4", "~5", "~6"})

	urc := sheet.MustGet("UsedRange", "Rows", "Count").(int32)
	println("str:"+sheet.MustCells(1, 2), sheet.MustGetCell(1, 2).(float64), cell.MustGet().(string), urc)

	cnt := 0
	sheet.ReadRow("A", 1, "F", 9, func(row []interface{}) (rc int) { //"A", 1 or 1, 9 or 1 or nothing
		cnt++
		fmt.Println(cnt, row)
		return //-1: break
	})

	time.Sleep(2000000000)

	//Sort
	cells := excel.GetIDispatch(sheet, "Cells")
	cells.CallMethod("UnMerge")
	sort := excel.GetIDispatch(sheet, "Sort")
	sortfields := excel.GetIDispatch(sort, "SortFields")
	sortfields.CallMethod("Clear")
	sortfields.CallMethod("Add", sheet.Range("f:f").IDispatch, 0, 2)
	sort.CallMethod("SetRange", cells)
	sort.PutProperty("Header", 1)
	sort.CallMethod("Apply")

	//Chart
	shapes := excel.GetIDispatch(sheet, "Shapes")
	_chart, _ := shapes.CallMethod("AddChart", 65)
	chart := _chart.ToIDispatch()
	chart.CallMethod("SetSourceData", sheet.Range("a1:c3").IDispatch)

	//AutoFilter
	cells.CallMethod("AutoFilter")
	excel.Release(sortfields, sort, cells, chart, shapes)

	time.Sleep(3000000000)
	xl.SaveAs(sNewPath+"/test_excel", "xls") //xl.SaveAs("test_excel", "html")

	xl.Quit()
	sheet.Release()
	cell.Release()
	rg.Release()

	fmt.Print("Press 'Enter' to continue...")
	bufio.NewReader(os.Stdin).ReadBytes('\n')
}
