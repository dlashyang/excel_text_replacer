/*
Copyright © 2022 NAME HERE <EMAIL ADDRESS>

*/
package main

import (
	"encoding/json"
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "excel_text_replacer src_file dst_file",
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
		srcName := args[0]
		var dstName string
		if len(args) == 2 {
			dstName = args[1]
		} else {
			dstName = args[0] + ".json"
		}
		fmt.Printf("coverting: %s -> %s\n", srcName, dstName)
		if err := convertFile(srcName, dstName); err != nil {
			log.Fatal(err)
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
	rootCmd.Flags().BoolP("no-new-file", "n", false, "save on the original file")
}

func main() {
	execute()
}

func convertFile(srcFileName, dstFileName string) error {
	f, err := excelize.OpenFile(srcFileName)
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

	mapD := make(map[string]map[string]string)
	for _, sheetName := range f.GetSheetList() {
		fmt.Println(sheetName)
		mapD[sheetName] = make(map[string]string)
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
				fmt.Print(rowCell, "\n")
				mapD[sheetName][coord] = rowCell
			}
			fmt.Print("\n\n")
		}
	}

	bolB, _ := json.MarshalIndent(mapD, "", "  ")
	fmt.Println(string(bolB))
	err = writeFile(dstFileName, bolB)
	if err != nil {
		log.Fatal(err)
		return err
	}

	return nil
}

func writeFile(fileName string, content []byte) error {
	filePtr, err := os.Create(fileName)
	if err != nil {
		log.Fatal(err)
		return err
	}
	defer filePtr.Close()

	_, err = filePtr.Write(content)
	if err != nil {
		log.Fatal(err)
		return err
	}

	return nil
}
