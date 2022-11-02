/*
Copyright © 2022 NAME HERE <EMAIL ADDRESS>
*/
package main

import (
	"bufio"
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
)

const h1 = "#  "
const h2 = "## "
const block = "'''"

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "excel_text_replacer excel_file_name",
	Short: "convert an excel file to text-format file",
	Long: `A longer description that spans multiple lines and likely contains
examples and usage of using your application. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	Args: cobra.MinimumNArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		flagUpdate, _ := cmd.Flags().GetBool("update")
		var excelFileName, textFileName string
		if len(args) == 2 {
			excelFileName = args[0]
			textFileName = args[1]
		} else {
			if flagUpdate {
				textFileName = args[0]
				excelFileName = strings.TrimSuffix(textFileName, ".md")
			} else {
				excelFileName = args[0]
				textFileName = excelFileName + ".md"
			}
		}

		if flagUpdate {
			fmt.Printf("updating: %s -> %s\n", textFileName, excelFileName)
			if err := text2excel(excelFileName, textFileName); err != nil {
				log.Fatal(err)
			}
		} else {
			fmt.Printf("coverting: %s -> %s\n", excelFileName, textFileName)
			if err := excel2text(excelFileName, textFileName); err != nil {
				log.Fatal(err)
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
	rootCmd.Flags().BoolP("update", "u", false, "update from text file")
}

func main() {
	execute()
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
