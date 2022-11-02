/*
Copyright © 2022 Leopold Yang
*/
package main

import (
	"bufio"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
)

const h1 = "#  "
const h2 = "## "
const block = "'''"
const helpMsg = `Usage:
  Converting to text file: excel_text_replacer excel_file text_file
  Updating excel file from text file: excel_text_replacer text_file excel_file`

func main() {
	flag.Parse()

	args := flag.Args()
	genName := false
	if len(args) < 1 {
		fmt.Println(helpMsg)
		os.Exit(0)
	} else if len(args) == 1 {
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
			os.Exit(1)
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
			os.Exit(1)
		}
	}
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

	fpText, err := os.Open(textFile)
	if err != nil {
		return fmt.Errorf("open text file fail: %s", err)
	}
	defer fpText.Close()

	var sheet, coord, newCell string
	scanner := bufio.NewScanner(fpText)
	for scanner.Scan() {
		line := scanner.Text()

		if strings.HasPrefix(line, h1) {
			sheet = strings.TrimPrefix(line, h1)
		} else if strings.HasPrefix(line, h2) {
			coord = strings.TrimPrefix(line, h2)
		} else {
			if strings.HasPrefix(line, block) {
				line = strings.TrimPrefix(line, block)
				newCell = ""
			}

			if strings.HasSuffix(line, block) {
				newCell += strings.TrimSuffix(line, block)
				//check or update cell
				cell, err := fpExcel.GetCellValue(sheet, coord)
				if err != nil {
					log.Fatal(err)
				}
				if newCell != cell {
					err = fpExcel.SetCellStr(sheet, coord, newCell)
					if err != nil {
						log.Fatal("cell update fail", sheet, coord, err)
						continue
					}
					log.Printf("Cell update: %s:%s\n-----\n%s\n=====\n%s\n++++\n", sheet, coord, cell, newCell)
				}
				continue
			}

			newCell += line + "\n"
		}
	}

	if err := scanner.Err(); err != nil {
		log.Fatal("scanner error", err)
	}

	if err = fpExcel.SaveAs("new_" + excelFile); err != nil {
		return fmt.Errorf("excel file save fail: %s", err)
	}

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

	textOut := ""
	for _, sheet := range fpExcel.GetSheetList() {
		log.Println("found sheet: ", sheet)
		textOut += h1 + sheet + "\n\n"
		// Get all the cols.
		cols, err := fpExcel.GetCols(sheet)
		if err != nil {
			return fmt.Errorf("excel get col fail: %s", err)
		}
		for i, col := range cols {
			for j, rowCell := range col {
				if rowCell == "" {
					continue
				}
				coord, err := excelize.CoordinatesToCellName(i+1, j+1)
				if err != nil {
					log.Fatal(err)
				}
				textOut += h2 + coord + "\n\n"
				textOut += block + rowCell + block + "\n\n"
				log.Printf("found cell %s:\n%s\n", coord, rowCell)
			}
		}
	}

	//fmt.Println(textOut)
	err = writeFile(textFile, textOut)
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
