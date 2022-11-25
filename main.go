/*
Copyright © 2022 Leopold Yang
*/
package main

import (
	"bufio"
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"os"
	"strings"
	"time"
)

const h1 = "#  "
const h2 = "## "
const block = "'''"

var flagDbgMsg bool
var filterSheet string
var dbgLog *log.Logger

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "excel_text_tool excel_file_name",
	Short: "convert an excel file to text-format file",
	Long: `A longer description that spans multiple lines and likely contains
examples and usage of using your application. For example:

  Converting to text file: excel_text_replacer excel_file text_file
  Updating excel file from text file: excel_text_replacer text_file excel_file`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	Args: cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		flagDbgMsg, _ = cmd.Flags().GetBool("verbos")
		filterSheet, _ = cmd.Flags().GetString("sheet")

		initLogger()
		genName := false
		if len(args) == 1 {
			genName = true
		}

		var excelFile, textFile string
		if strings.HasSuffix(args[0], ".xlsx") {
			excelFile = args[0]
			if genName {
				textFile = excelFile + ".md"
			} else {
				textFile = args[1]
			}
			log.Printf("coverting: %s -> %s\n", excelFile, textFile)
			if err := excel2text(excelFile, textFile); err != nil {
				log.Fatal("coverting excel to text fail: ", err)
			}
		} else {
			textFile = args[0]
			if genName {
				excelFile = strings.TrimSuffix(textFile, ".md")
			} else {
				excelFile = args[1]
			}
			log.Printf("updating: %s -> new_%s\n", textFile, excelFile)
			if err := text2excel(excelFile, textFile); err != nil {
				log.Fatal("updating excel from text fail: ", err)
			}
		}
	},
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	// Here you will define your flags and configuration settings.
	// Cobra supports persistent flags, which, if defined here,
	// will be global for your application.

	// rootCmd.PersistentFlags().StringVar(&cfgFile, "config", "", "config file (default is $HOME/.excel_text_replacer.yaml)")

	// Cobra also supports local flags, which will only run
	// when this action is called directly.
	rootCmd.Flags().BoolP("verbos", "v", false, "print debug info")
	rootCmd.Flags().String("sheet", "", "working on given sheet only")
}

func initLogger() {
	file, err := os.OpenFile("log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0666)
	if err != nil {
		log.Fatal(err)
	}
	mw := io.MultiWriter(os.Stdout, file)
	log.SetOutput(mw)

	dbgLog = log.New(file, "DEBUG: ", log.Ltime|log.Lmicroseconds|log.Lshortfile)
	if !flagDbgMsg {
		dbgLog.SetOutput(io.Discard)
	}
}

func main() {
	start := time.Now()

	execute()

	log.Println("done: ", time.Since(start))
}

func text2excel(excelFile, textFile string) error {
	fpExcel, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel file fail: %s", err)
	}
	defer func() {
		if err := fpExcel.Close(); err != nil {
			log.Fatal(err)
		}
	}()
	dbgLog.Println("open excel file successfully")

	fpText, err := os.Open(textFile)
	if err != nil {
		return fmt.Errorf("open text file fail: %s", err)
	}
	defer fpText.Close()
	dbgLog.Println("open text file successfully")

	var sheet, coord string
	var b strings.Builder
	cellUpdated := 0
	scanner := bufio.NewScanner(fpText)
	flagBlockStart := false
	for scanner.Scan() {
		line := scanner.Text()
		dbgLog.Printf("reading line:\n%s", line)

		if strings.HasPrefix(line, h1) {
			sheet = strings.TrimPrefix(line, h1)
			dbgLog.Println("found sheet: ", sheet)
		} else if strings.HasPrefix(line, h2) {
			coord = strings.TrimPrefix(line, h2)
			dbgLog.Println("found cell: ", coord)
		} else {
			if !flagBlockStart {
				dbgLog.Println("flagBlockStart is off")
				if strings.HasPrefix(line, block) {
					line = strings.TrimPrefix(line, block)
					b.Reset()
					flagBlockStart = true
					dbgLog.Println("block starts")
				}
			}

			if strings.HasSuffix(line, block) {
				b.WriteString(strings.TrimSuffix(line, block))
				flagBlockStart = false
				newCell := b.String()
				dbgLog.Printf("block ends\n%s", newCell)

				//check or update cell
				cell, err := fpExcel.GetCellValue(sheet, coord)
				if err != nil {
					log.Fatal(err)
				}
				dbgLog.Printf("read cell %s:%s from excel\n%s", sheet, coord, cell)

				if newCell != cell {
					err = fpExcel.SetCellStr(sheet, coord, newCell)
					if err != nil {
						log.Println("cell update fail:", sheet, coord, err)
						continue
					}
					cellUpdated++
					log.Printf("cell update: %s:%s\n", sheet, coord)
					dbgLog.Printf("cell update: %s:%s\n-----\n%s\n=====\n%s\n++++\n", sheet, coord, cell, newCell)
				}
				continue
			}

			b.WriteString(line + "\n")
		}
	}

	if err := scanner.Err(); err != nil {
		log.Fatal("scanner error", err)
	}

	if err = fpExcel.SaveAs("new_" + excelFile); err != nil {
		return fmt.Errorf("excel file save fail: %s", err)
	}
	log.Println("cell updated: ", cellUpdated)

	return nil
}

func excel2text(excelFile, textFile string) error {
	fpExcel, err := excelize.OpenFile(excelFile)
	if err != nil {
		return fmt.Errorf("open excel fail: %s", err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := fpExcel.Close(); err != nil {
			log.Fatal(err)
		}
	}()
	dbgLog.Println("open excel file successfully")

	var bTextOut strings.Builder
	cellFound := 0
	for _, sheet := range fpExcel.GetSheetList() {
		dbgLog.Printf("found sheet %s with filter [%s]", sheet, filterSheet)
		if (filterSheet != "") && (filterSheet != sheet) {
			continue
		}

		log.Println("found sheet: ", sheet)
		bTextOut.WriteString(h1 + sheet + "\n\n")
		// Get all the cols.
		rows, err := fpExcel.GetRows(sheet)
		if err != nil {
			return fmt.Errorf("excel get rows fail: %s", err)
		}
		for i, row := range rows {
			for j, cell := range row {
				if cell == "" {
					continue
				}
				coord, err := excelize.CoordinatesToCellName(j+1, i+1)
				if err != nil {
					log.Fatal(err)
				}
				bTextOut.WriteString(h2 + coord + "\n\n" + block + cell + block + "\n\n")
				cellFound++
				dbgLog.Printf("found cell %s:%s", sheet, coord)
			}
		}
	}

	log.Println("Found cell: ", cellFound)
	err = writeFile(textFile, bTextOut.String())
	if err != nil {
		return fmt.Errorf("write text file fail: %s", err)
	}

	return nil
}

func writeFile(fileName string, content string) error {
	filePtr, err := os.Create(fileName)
	if err != nil {
		return fmt.Errorf("create text file fail: %s", err)
	}
	defer filePtr.Close()

	_, err = filePtr.WriteString(content)
	if err != nil {
		return fmt.Errorf("file write error: %s", err)
	}

	return nil
}
