package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"mswordrefiner/utils"
)

func main() {

	if len(os.Args) < 2 {
		displayUsage()
		return
	}

	initialWordLocation := os.Args[1]

	ext := filepath.Ext(initialWordLocation)
	baseName := strings.TrimSuffix(initialWordLocation, ext)

	newFilename := fmt.Sprintf("%s_c%s", baseName, ext)

	docxContent, err := utils.LoadDocxFile(initialWordLocation)
	if err != nil {
		fmt.Printf("Loading .docx file failed: %s\n", err)
		return
	}

	// Now using the modified function that returns a new zip.Reader
	modifiedContent, err := utils.RemoveCommentsFromDocument(docxContent)
	if err != nil {
		fmt.Printf("Error occured while remocing comments: %s\n", err)
		return
	}

	err = utils.WriteModifiedDocxToDisk(modifiedContent, newFilename)
	if err != nil {
		fmt.Printf("Failed to write file to file system: %s\n", err)
		return
	}

	fmt.Println("Done!")
	fmt.Printf("New file: %s\n", newFilename)

}

func displayUsage() {
	fmt.Println("")
	fmt.Println("MS Word Refiner v.0.1")
	fmt.Println("")
	fmt.Println("Removes comments and reviews from a DOCX file. Usefull for Mac OS where document inspection is not possible.")
	fmt.Println("The new file will be saved in the same folder as the original file, with _c suffix.")
	fmt.Println("")
	fmt.Println("Usage:")
	fmt.Printf("  %s <input_docx_file>\n", os.Args[0])
}
