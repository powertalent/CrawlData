package main

import (
	"bufio"
	"fmt"
	"os"
	"path"
	"regexp"
	"strings"

	"github.com/gocolly/colly/v2"
	"github.com/xuri/excelize/v2"
)

func readURLsFromFile(filePath string) ([]string, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	var urls []string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		urls = append(urls, scanner.Text())
	}

	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return urls, nil
}

func getColumnName(col int) string {
	name := make([]byte, 0, 3) // max 16,384 columns (2022)
	const aLen = 'Z' - 'A' + 1 // alphabet length
	for ; col > 0; col /= aLen + 1 {
		name = append(name, byte('A'+(col-1)%aLen))
	}
	for i, j := 0, len(name)-1; i < j; i, j = i+1, j-1 {
		name[i], name[j] = name[j], name[i]
	}
	return string(name)
}

func cleanSheetName(name string) string {
	// Characters not allowed in Excel sheet names
	invalidChars := []string{":", "\\", "/", "?", "*", "[", "]"}

	// Replace each invalid character with an underscore or remove it
	for _, char := range invalidChars {
		name = strings.Replace(name, char, "", -1) // removes the character
		// name = strings.Replace(name, char, "_", -1) // or replace it with an underscore
	}

	// Ensure the sheet name is not too long (Excel limit is 31 characters)
	if len(name) > 31 {
		name = name[:31]
	}

	return name
}

func main() {
	urls, err := readURLsFromFile("links.txt")
	if err != nil {
		fmt.Println("Error reading URLs from file:", err)
		return
	}

	// Create a new Excel file
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Counter for Excel rows
	row := 1

	// Instantiate default collector
	c := colly.NewCollector(
		colly.AllowedDomains("www.supermicro.com"),
	)

	colIdx := 1
	// On every a element which has href attribute call callback
	c.OnHTML("[class^=specHeader]", func(e *colly.HTMLElement) {
		header := e.Text
		// Print link
		fmt.Printf("Link found: %q -> %s\n", e.Text, header)
		parentTableDOM := e.DOM.Closest("table")

		// // Use goquery to manipulate HTML
		// doc := goquery.NewDocumentFromNode(parentTableDOM.Get(0))

		// // Find all <ul> elements and replace them with their own contents
		// doc.Find("ul").Each(func(index int, item *goquery.Selection) {
		// 	contentHtml, _ := item.Html()
		// 	item.ReplaceWithHtml(contentHtml)
		// })

		// Use goquery to manipulate HTML
		parentTableDOM.Find("tr:has([class^=specHeader])").Prev().Remove() // This removes the row before header (Empty row)
		parentTableDOM.Find("img").Closest("tr").Remove()                  // This removes all <img> tags within the table

		parentTableHtml, err := parentTableDOM.Html()
		if err != nil {
			fmt.Println("Error getting parent HTML:", err)
			return
		}

		// Regex to find <ul> tags and their contents
		re := regexp.MustCompile(`<ul.*?>|</ul>`)
		parentTableHtml = re.ReplaceAllString(parentTableHtml, "")

		parentTableHtml = "<table>" + parentTableHtml + "</table>"

		// Logic to determine the sheet name based on header content
		var sheetName string
		if strings.Contains(header, "Product SKUs") {
			sheetName = "Product SKUs"
		} else if strings.Contains(header, "Processor") {
			sheetName = "Processor"
		} else if strings.Contains(header, "Chassis") {
			sheetName = "Chassis"
		} else {
			sheetName = cleanSheetName(header)
		}

		// Get the current list of sheet names
		sheetMap := f.GetSheetMap()
		exists := false
		var sheetIndex int

		// Check if the sheet already exists
		for idx, name := range sheetMap {
			if name == sheetName {
				exists = true
				sheetIndex = idx
				break
			}
		}

		if !exists {
			// Create a new sheet if it does not exist
			sheetIndex, err = f.NewSheet(sheetName)
			if err != nil {
				fmt.Println("Error creating new sheet:", err)
				return
			}
		} else {
			// Set the existing sheet as the active sheet if it exists
			f.SetActiveSheet(sheetIndex)
		}

		// Write data to Excel
		cell := fmt.Sprintf("%s%d", getColumnName(colIdx), row)
		f.SetCellValue(sheetName, cell, parentTableHtml)

	})

	// Before making a request print "Visiting ..."
	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL.String())

	})

	// Loop through all URLs from the file
	cnt := 1
	for _, url := range urls {
		fmt.Println("baseURL ", path.Base(url))
		f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", getColumnName(colIdx), 1), cnt)              // index
		f.SetCellValue("Sheet1", fmt.Sprintf("%s%d", getColumnName(colIdx+1), 1), path.Base(url)) // index
		c.Visit(url)
		cnt++
		colIdx += 2
	}

	// Set active sheet of the workbook
	f.SetActiveSheet(1)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("mainboardList.xlsx"); err != nil {
		fmt.Println(err)
	}
}
