/*
Copyright © 2022 NAME HERE <EMAIL ADDRESS>

*/
package main

import (
	"fmt"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "excel_text_replacer old_text new_text target_file",
	Short: "A brief description of your application",
	Long: `A longer description that spans multiple lines and likely contains
examples and usage of using your application. For example:

Cobra is a CLI library for Go that empowers applications.
This application is a tool to generate the needed files
to quickly create a Cobra application.`,
	// Uncomment the following line if your bare application
	// has an action associated with it:
	Args: cobra.MinimumNArgs(3),
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Printf("excuting: [%s] %s -> %s\n", args[2], args[0], args[1])
		if err := replaceinFile(args[0], args[1], args[2]); err != nil {
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
	// execute()
	fmt.Println(os.Args[1])
	fmt.Println(restoreEscChar(os.Args[1]))
}

func replaceinFile(oldText, newText, fileName string) error {
	fmt.Println(oldText)
	restoredText := restoreEscChar(oldText)
	f, err := excelize.OpenFile(fileName)
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

	// Get all the cols in the Sheet1.
	cols, err := f.GetCols("Sheet2")
	if err != nil {
		log.Fatal(err)
		return err
	}
	for _, col := range cols {
		for _, rowCell := range col {
			// fmt.Print(rowCell, "\n")
			// fmt.Print("\n", oldText, "\n")
			newCell := strings.Replace(rowCell, restoredText, newText, -1)
			fmt.Print("\n", newCell, "\n")
		}
		fmt.Print("\n\n")
	}

	return nil
}

func restoreEscChar(oriString string) string {
	// test escape character
	escCharMap := map[string]string{"\\n": "\n", "\\t": "\t"}

	printChar(oriString)

	expectedString := oriString
	for k, v := range escCharMap {
		//fmt.Printf("%s: %s.\n", k, v)
		expectedString = strings.ReplaceAll(expectedString, k, v)
	}

	printChar(expectedString)

	return expectedString
}

func printChar(str string) {
	char := []byte(str)
	fmt.Println("len:", len(char))
	for _, c := range char {
		fmt.Printf("%c[%d] ", c, c)
	}
	fmt.Println()
}
