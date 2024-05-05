package main

import (
	"bufio"
	"fmt"
	"os"

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

func main() {

	// Create a new Excel file
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Counter for Excel rows
	row := 2

	// Instantiate default collector
	c := colly.NewCollector(
		colly.AllowedDomains("www.supermicro.com"),
	)

	// On every a element which has href attribute call callback
	c.OnHTML(".scrollable-table [class^=specHeader]", func(e *colly.HTMLElement) {
		header := e.Text
		// Print link
		fmt.Printf("Link found: %q -> %s\n", e.Text, header)
		parentTableDOM := e.DOM.Closest("table").Parent()
		parentTableHtml, err := parentTableDOM.Html()
		if err != nil {
			fmt.Println("Error getting parent HTML:", err)
			return
		}

		// Create a new sheet
		sheetName := header
		index, err := f.NewSheet(sheetName)
		if err != nil {
			fmt.Println(err)
			return
		}
		// Set value of a cell
		f.SetCellValue(sheetName, "A1", "Link Text")
		f.SetCellValue(sheetName, "B1", "URL")

		// Write data to Excel
		cellA := fmt.Sprintf("A%d", row)
		cellB := fmt.Sprintf("B%d", row)
		f.SetCellValue(sheetName, cellA, e.Text)
		f.SetCellValue(sheetName, cellB, header)
		// parentHtml, err := e.DOM.Parent().Parent().Parent().Parent().Html() // Capture both the HTML and the error
		parentHtml, err := e.DOM.Closest("table").Html()
		if err != nil {
			fmt.Println("Error getting parent HTML:", err)
			return
		}
		f.SetCellValue("Sheet1", "C1", parentHtml)
		// Visit link found on page
		// Only those links are visited which are in AllowedDomains
		// c.Visit(e.Request.AbsoluteURL(link))
	})

	// Before making a request print "Visiting ..."
	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL.String())

	})

	c.Visit("https://www.supermicro.com/ja/products/motherboard/A1SAM-2750F")

	// Set active sheet of the workbook
	f.SetActiveSheet(1)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
