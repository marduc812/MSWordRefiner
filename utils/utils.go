package utils

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"regexp"
	"slices"
	"strings"
)

func LoadDocxFile(selectedFile string) (*zip.Reader, error) {
	fileBytes, err := os.ReadFile(selectedFile)
	if err != nil {
		return nil, err
	}

	reader := bytes.NewReader(fileBytes)
	docxReader, err := zip.NewReader(reader, int64(len(fileBytes)))
	if err != nil {
		return nil, err
	}

	return docxReader, nil
}

// RemoveCommentsFromDocument takes a zip.Reader representing a DOCX file and removes specific comment-related XML files.
// It returns a new zip.Reader with the modified content or an error if any occurs.
//
// Parameters:
//   - docxFile: A zip.Reader containing the input DOCX file's content.
//
// Returns:
//   - *zip.Reader: A new zip.Reader with the modified content after removing specified comment files.
//   - error: An error if any issues occur during the process.
func RemoveCommentsFromDocument(docxFile *zip.Reader) (*zip.Reader, error) {
	documentLocation := "word/document.xml"
	settingsLocation := "word/settings.xml"
	commentFiles := []string{"word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsIdsExt.xml", "word/commentsExtensible.xml"}

	newDocxData := new(bytes.Buffer)
	docxWriter := zip.NewWriter(newDocxData)

	for _, f := range docxFile.File {

		var docxContentBytes []byte

		rc, err := f.Open()
		if err != nil {
			fmt.Println(err)
			return nil, err
		}

		content, err := io.ReadAll(rc)
		rc.Close()

		if err != nil {
			fmt.Println(err)
			return nil, err
		}

		if slices.Contains(commentFiles, f.Name) {
			continue
		}

		// Remove comments from the document.xml file
		if f.Name == documentLocation {
			removeCommentRange := regexp.MustCompile(`<w:commentRangeStart [^>]*\/>.*?<w:commentRangeEnd [^>]*\/>`)
			removeCommentReference := regexp.MustCompile(`<w:commentReference [^>]*\/>`)

			contentString := string(content)

			contentString = removeCommentRange.ReplaceAllString(contentString, "")
			contentString = removeCommentReference.ReplaceAllString(contentString, "")

			regexDeletedChange := regexp.MustCompile(`<w:del [^>]*>.*?<\/w:del>`)
			regexInsertedChange := regexp.MustCompile(`<w:ins [^>]*>(.*?)<\/w:ins>`)
			regexFormatChange := regexp.MustCompile(`<w:rPrChange [^>]*>.*?<\/w:rPrChange>`)

			contentString = regexDeletedChange.ReplaceAllString(contentString, "")
			contentString = regexInsertedChange.ReplaceAllString(contentString, "$1")
			contentString = regexFormatChange.ReplaceAllString(contentString, "")

			docxContentBytes = []byte(contentString)
		} else if f.Name == settingsLocation {
			// disable track changes
			contentString := string(content)
			contentString = strings.ReplaceAll(contentString, "<w:trackRevisions/>", "")
			docxContentBytes = []byte(contentString)

		} else {
			docxContentBytes = content
		}

		fw, err := docxWriter.Create(f.Name)
		if err != nil {
			fmt.Println(err)
			return nil, err
		}

		_, err = fw.Write(docxContentBytes)
		if err != nil {
			fmt.Println(err)
			return nil, err
		}
	}

	err := docxWriter.Close()
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	readerAt := bytes.NewReader(newDocxData.Bytes())
	newDocxReader, err := zip.NewReader(readerAt, int64(len(newDocxData.Bytes())))
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	return newDocxReader, nil

}

func WriteModifiedDocxToDisk(docxContent *zip.Reader, outputPath string) error {
	// Create a new file on the disk to write the modified docx content
	newFile, err := os.Create(outputPath)
	if err != nil {
		return err
	}
	defer newFile.Close()

	// Create a new zip writer on this file
	archive := zip.NewWriter(newFile)
	defer archive.Close()

	// Iterate through each file in the original docx content
	for _, file := range docxContent.File {
		header, err := zip.FileInfoHeader(file.FileInfo())
		if err != nil {
			return err
		}
		header.Name = file.Name
		header.Method = zip.Deflate

		writer, err := archive.CreateHeader(header)
		if err != nil {
			return err
		}

		reader, err := file.Open()
		if err != nil {
			return err
		}

		// If it's the document.xml (or any other modified file), you'll replace this with your modified content
		if file.Name == "word/document.xml" {
			// Example: Replace with modified content
			// writer.Write(yourModifiedContent)
			// For now, just copying the original content
			io.Copy(writer, reader)
		} else {
			// If not, just copy the original file into the archive
			io.Copy(writer, reader)
		}
		reader.Close()
	}

	return nil
}
