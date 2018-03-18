package main

import (
	"github.com/tealeg/xlsx"
	"github.com/sqweek/dialog"
	"github.com/andlabs/ui"
	"path/filepath"
	"log"
	"github.com/extrame/xls"
	"strconv"
	"path"
	"strings"
	"os"
	"sync"
	"time"
)

var directory = ""
var errors = 0
var success = 0
var mutexError = &sync.Mutex{}
var mutexSuccess = &sync.Mutex{}

var Log *log.Logger

func selectMainDir() {
	directoryPath, err := dialog.Directory().Title("Select excel directory").Browse()
	if err != nil {
		Log.Println(err)
		os.Exit(0)
	}
	directory = directoryPath
}

func listDir() []string {
	files, err := filepath.Glob(filepath.Join(directory, "*.xls"))
	if err != nil {
		Log.Fatal(err)
	}
	return files // contains a list of all files in the targeted director
}


func printAlertResult(totalCount int, alertMessage *ui.Label, stateMessage *ui.Label) {
	successCount, errorsCount := getResultat()
	alertMessage.SetText(strconv.Itoa(successCount) + " Files converted, " + strconv.Itoa(errorsCount) + " errors.")
	if successCount + errorsCount < totalCount {
		stateMessage.SetText("On going...")
		time.Sleep(100 * time.Millisecond)
		printAlertResult(totalCount, alertMessage, stateMessage)
	} else {
		stateMessage.SetText("Done! If errors were encountered, check converter.log")
	}
}

func resultWindow(window *ui.Window, fileNumberStr string) {
	window.SetChild(nil)
	box := ui.NewVerticalBox()
	box.Append(ui.NewLabel("Converting " + fileNumberStr + " files ..."), false)
	alertMessage := ui.NewLabel("")
	stateMessage := ui.NewLabel("")
	box.Append(alertMessage, false)
	box.Append(stateMessage, false)
	window.SetChild(box)
	totalCount, _ := strconv.Atoi(fileNumberStr)
	printAlertResult(totalCount, alertMessage, stateMessage)
}

func mainWindow(window *ui.Window) {
	window.SetChild(nil)
	buttonOk := ui.NewButton("Convert")
	buttonCancel := ui.NewButton("Cancel")
	box := ui.NewVerticalBox()
	fileNumberStr := strconv.Itoa(len(listDir()))
	box.Append(ui.NewLabel("You will convert " + fileNumberStr + " files in the directory "+ directory + ". Are you sure?"), false)
	box.Append(ui.NewLabel(""), false)
	box.Append(buttonOk, false)
	box.Append(buttonCancel, false)
	buttonCancel.OnClicked(func(*ui.Button) {
		selectMainDir()
		mainWindow(window)
	})
	buttonOk.OnClicked(func(*ui.Button) {
		go convertDirectoryToXlsx(false)
		resultWindow(window, fileNumberStr)
	})
	window.SetChild(box)
}

func saveXlsx(xlsxFile *xlsx.File, excelFileName string) {
	err := xlsxFile.Save(excelFileName)
	if err != nil {
		Log.Println(err.Error())
		incResultat(false)
	} else {
		incResultat(true)
	}
}

func convertXlsToXlsx(xlsPath string, removeFile bool) {
	var xlsxSheet *xlsx.Sheet

	pathFile := strings.TrimSuffix(xlsPath, filepath.Ext(xlsPath))
	xlsxName := path.Join(pathFile + ".xlsx")

	xlsxFile := xlsx.NewFile()
	defer saveXlsx(xlsxFile, xlsxName)

	if xlFile, err := xls.Open(xlsPath, "utf-8"); err == nil {

		for sheet_index := 0; sheet_index <= xlFile.NumSheets(); sheet_index ++ {

			if xlsSheet := xlFile.GetSheet(sheet_index); xlsSheet != nil {
				// add new sheet
				xlsxSheet, err = xlsxFile.AddSheet(xlsSheet.Name)

				// iteration each row
				for row_index := 0; row_index <= (int(xlsSheet.MaxRow)); row_index++ {
					// create corresponding row
					xlsRow := xlsSheet.Row(row_index)
					xlsxRow := xlsxSheet.AddRow()

					// iteration each column
					for col_index := 0; col_index <=  xlsRow.LastCol(); col_index ++ {
						// create and fill corresponding col
						cell := xlsxRow.AddCell()
						cell.Value = xlsRow.Col(col_index)
					}
				}
			}
		}
		xlFile = nil
	} else {
		Log.Println(err)
		incResultat(false)
		return
	}
	if removeFile {
		err := os.Remove(xlsPath)
		if err != nil {
			Log.Println(err)
		}
	}
}

func incResultat(state bool) {
	if state {
		mutexSuccess.Lock()
		success += 1
		mutexSuccess.Unlock()
	} else {
		mutexError.Lock()
		errors += 1
		mutexError.Unlock()
	}
}

func getResultat() (int, int) {
	mutexSuccess.Lock()
	mutexError.Lock()
	success_ret := success
	errors_ret := errors
	mutexSuccess.Unlock()
	mutexError.Unlock()
	return  success_ret, errors_ret
}

func convertDirectoryToXlsx(removeFile bool) {
	for _, xlsPath := range(listDir()) {
		convertXlsToXlsx(xlsPath, removeFile)
	}
}

func main() {

	file, errLog := os.OpenFile("converter.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)

	if errLog != nil {
		panic(errLog)
	}

	Log = log.New(file, "", log.LstdFlags|log.Lshortfile)

	selectMainDir()

	err := ui.Main(func() {

		window := ui.NewWindow("Conversion confirmation", 500, 200, false)
		window.SetMargined(true)
		mainWindow(window)

		window.OnClosing(func(*ui.Window) bool {
			ui.Quit()
			return true
		})
		window.Show()
	})
	if err != nil {
		panic(err)
	}
}
