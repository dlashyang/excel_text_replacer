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
		return
	} else if len(args) == 1 {
		genName = true
	}

	var excelFileName, textFileName string
	if strings.HasSuffix(args[0], ".xlsx") {
		excelFileName = args[0]
		if genName {
			textFileName = excelFileName + ".md"
		} else {
			textFileName = args[1]
		}
		fmt.Printf("coverting: %s -> %s\n", excelFileName, textFileName)
		if err := excel2text(excelFileName, textFileName); err != nil {
			log.Fatal(err)
		}
	} else {
		textFileName := args[0]
		if genName {
			excelFileName = strings.TrimSuffix(textFileName, ".md")
		} else {
			excelFileName = args[1]
		}
		fmt.Printf("updating: %s -> new_%s\n", textFileName, excelFileName)
		if err := text2excel(excelFileName, textFileName); err != nil {
			log.Fatal(err)
		}
	}
}

func text2excel(excelFile, textFile string) error {
	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		log.Fatal(err)
		return err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	file, err := os.Open(textFile)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	var sheetName, coord, cellContent string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		content := scanner.Text()
		//fmt.Println("line 100: ", content)

		if strings.HasPrefix(content, h1) {
			sheetName = strings.TrimPrefix(content, h1)
			//fmt.Println("sheetname: ", sheetName)
		} else if strings.HasPrefix(content, h2) {
			coord = strings.TrimPrefix(content, h2)
			//fmt.Println("coordination: ", coord)
		} else {
			if strings.HasPrefix(content, block) {
				content = strings.TrimPrefix(content, block)
				cellContent = ""
			}

			if strings.HasSuffix(content, block) {
				cellContent += strings.TrimSuffix(content, block)
				//check or update cell
				cell, err := f.GetCellValue(sheetName, coord)
				if err != nil {
					log.Fatal(err)
					return (err)
				}
				if cellContent != cell {
					f.SetCellStr(sheetName, coord, cellContent)
					fmt.Println("Cell update: ", sheetName, coord)
					fmt.Println(cell, " -> ", cellContent)
				}
				continue
			}

			cellContent += content + "\n"
			//fmt.Println("content is: \n", content)
		}
	}

	if err := scanner.Err(); err != nil {
		log.Fatal(err)
	}

	if err = f.SaveAs("new_" + excelFile); err != nil {
		log.Fatal(err)
	}

	return nil
}

func excel2text(excelFile, textFile string) error {
	f, err := excelize.OpenFile(excelFile)
	if err != nil {
		log.Fatal(err)
		return err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	textOutput := ""
	for _, sheetName := range f.GetSheetList() {
		fmt.Println(sheetName)
		textOutput += h1 + sheetName + "\n\n"
		// Get all the cols.
		cols, err := f.GetCols(sheetName)
		if err != nil {
			log.Fatal(err)
			return err
		}
		for i, col := range cols {
			for j, rowCell := range col {
				if rowCell == "" {
					continue
				}
				coord, err := excelize.CoordinatesToCellName(i+1, j+1)
				if err != nil {
					log.Fatal(err)
					return err
				}
				fmt.Print("\n", coord, "\n")
				textOutput += h2 + coord + "\n\n"
				fmt.Print(rowCell, "\n")
				textOutput += block + rowCell + block + "\n\n"
			}
			fmt.Print("\n\n")
		}
	}

	fmt.Println(textOutput)
	err = writeFile(textFile, textOutput)
	if err != nil {
		log.Fatal(err)
		return err
	}

	return nil
}

func writeFile(fileName string, content string) error {
	filePtr, err := os.Create(fileName)
	if err != nil {
		log.Fatal(err)
		return err
	}
	defer filePtr.Close()

	_, err = filePtr.WriteString(content)
	if err != nil {
		log.Fatal(err)
		return err
	}

	return nil
}
