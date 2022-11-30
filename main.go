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

var filterSheet string
var dbgLog *log.Logger

// rootCmd represents the base command when called without any subcommands
var convertCmd = &cobra.Command{
	Use:   "convert excel_file",
	Short: "Convert excel file to text file",
	Long: `Convert excel file to text format(markdown) file. For example:

  excel_text_tool convert excel_file.xlsx
  excel_text_tool convert excel_file.xlsx -o target.md
  excel_text_tool convert excel_file.xlsx --sheet Sheet1`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	Args: cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		flagDbgMsg, _ := cmd.PersistentFlags().GetBool("verbos")
		initLogger(flagDbgMsg)

		filterSheet, _ = cmd.Flags().GetString("sheet")
		textFile, _ := cmd.Flags().GetString("output")

		excelFile := args[0]
		if textFile == "" {
			textFile = excelFile + ".md"
		}
		log.Printf("coverting: %s -> %s\n", excelFile, textFile)
		if err := excel2text(excelFile, textFile); err != nil {
			log.Fatal("coverting excel to text fail: ", err)
		}
	},
}

var updateCmd = &cobra.Command{
	Use:   "update excel_file -i text_file",
	Short: "Update excel file content based on specific text file",
	Long: `Update excel file based on info provided by specific text file. For example:

  excel_text_tool excel_file -i text_file[Mandantory]
  excel_text_tool excel_file -i text_file -o save_as_new_excel_file
  excel_text_tool excel_file -i text_file --overwrite`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	Args: cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		flagDbgMsg, _ := cmd.PersistentFlags().GetBool("verbos")
		initLogger(flagDbgMsg)

		textFile, _ := cmd.Flags().GetString("input")
		outputFile, _ := cmd.Flags().GetString("output")
		flagOverwrite, _ := cmd.Flags().GetBool("overwrite")

		excelFile := args[0]
		if flagOverwrite {
			outputFile = excelFile
		}

		if outputFile == "" {
			outputFile = "new_" + excelFile
		}

		log.Printf("updating: %s -> %s based on [%s]\n", excelFile, outputFile, textFile)

		if err := text2excel(excelFile, textFile, outputFile); err != nil {
			log.Fatal("updating excel from text fail: ", err)
		}
	},
}

var rootCmd = &cobra.Command{
	Use:   "excel_text_tool",
	Short: "Tools for text work on excel file",
	Args:  cobra.NoArgs,
}

func init() {
	// Here you will define your flags and configuration settings.
	// Cobra supports persistent flags, which, if defined here,
	// will be global for your application.

	// rootCmd.PersistentFlags().StringVar(&cfgFile, "config", "", "config file (default is $HOME/.excel_text_replacer.yaml)")

	// Cobra also supports local flags, which will only run
	// when this action is called directly.
	rootCmd.PersistentFlags().BoolP("verbos", "v", false, "print debug info")
	rootCmd.AddCommand(convertCmd)
	rootCmd.AddCommand(updateCmd)

	convertCmd.Flags().String("sheet", "", "convert specific sheet only")
	convertCmd.Flags().StringP("output", "o", "", "output file name")

	updateCmd.Flags().StringP("input", "i", "", "input text file")
	updateCmd.MarkFlagRequired("input")
	updateCmd.Flags().StringP("output", "o", "", "output file name")
	updateCmd.Flags().Bool("overwrite", false, "update original excel file")
	updateCmd.MarkFlagsMutuallyExclusive("output", "overwrite")
}

func initLogger(flagDbgMsg bool) {
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

	err := rootCmd.Execute()
	if err != nil {
		os.Exit(0)
	}

	log.Println("done: ", time.Since(start))
}

func text2excel(excelFile, textFile, outputFile string) error {
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

	if err = fpExcel.SaveAs(outputFile); err != nil {
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
