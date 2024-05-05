package main

import (
	"encoding/json"
	"fmt"
	"log"
	"regexp"
	"strings"

	"github.com/PuerkitoBio/goquery"
	"github.com/atotto/clipboard"
	"github.com/gocolly/colly/v2"
)

type SubCategory struct {
	SubCategoryName string `json:"SubCategory"`
	Details         string `json:"Details"`
}

type MainCategory struct {
	CategoryName  string        `json:"MainCategory"`
	SubCategories []SubCategory `json:"SubCategories"`
}

func main() {

	// Create a new Excel file
	// f := excelize.NewFile()
	// defer func() {
	// 	if err := f.Close(); err != nil {
	// 		fmt.Println(err)
	// 	}
	// }()
	// Create a new sheet
	// index, err := f.NewSheet("Sheet1")
	// if err != nil {
	// 	fmt.Println(err)
	// 	return
	// }
	// Set value of a cell
	// f.SetCellValue("Sheet1", "A1", "Link Text")
	// f.SetCellValue("Sheet1", "B1", "URL")

	// Counter for Excel rows
	// row := 2

	// Instantiate default collector
	c := colly.NewCollector(
		colly.AllowedDomains("www.supermicro.com"),
	)

	data := []MainCategory{}

	var re = regexp.MustCompile(`\s+`)

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

		// Convert data to JSON
		subCategories, err := processTable(parentTableHtml)
		if err != nil {
			log.Fatal(err)
		}

		mainCategory := MainCategory{
			CategoryName:  e.Text,
			SubCategories: subCategories,
		}

		data = append(data, mainCategory)

		// // Write data to Excel
		// cellA := fmt.Sprintf("A%d", row)
		// cellB := fmt.Sprintf("B%d", row)
		// f.SetCellValue("Sheet1", cellA, e.Text)
		// f.SetCellValue("Sheet1", cellB, header)
		// // parentHtml, err := e.DOM.Parent().Parent().Parent().Parent().Html() // Capture both the HTML and the error
		// parentHtml, err := e.DOM.Closest("table").Html()
		// if err != nil {
		// 	fmt.Println("Error getting parent HTML:", err)
		// 	return
		// }
		// f.SetCellValue("Sheet1", "C1", parentHtml)
		// Visit link found on page
		// Only those links are visited which are in AllowedDomains
		// c.Visit(e.Request.AbsoluteURL(link))
	})

	c.OnScraped(func(r *colly.Response) {
		// Convert to JSON with indentation
		jsonData, err := json.MarshalIndent(data, "", "    ")
		if err != nil {
			fmt.Println("Error marshaling to JSON:", err)
			return
		}

		jsonDataStr := re.ReplaceAllString(string(jsonData), " ")

		// fmt.Print(jsonDataStr)
		err = clipboard.WriteAll(jsonDataStr)
		if err != nil {
			fmt.Println("Failed to copy text to the clipboard:", err)
			return
		}
	})

	// Before making a request print "Visiting ..."
	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL.String())

	})

	c.Visit("https://www.supermicro.com/ja/products/motherboard/A1SAM-2750F")

	// Set active sheet of the workbook
	// f.SetActiveSheet(index)
	// // Save spreadsheet by the given path.
	// if err := f.SaveAs("Book1.xlsx"); err != nil {
	// 	fmt.Println(err)
	// }
}

// processTable parses the HTML and returns a slice of SubCategory
func processTable(html string) ([]SubCategory, error) {
	var data []SubCategory
	// Create a new reader from the HTML string
	reader := strings.NewReader(html)
	// Load the HTML document
	doc, err := goquery.NewDocumentFromReader(reader)
	if err != nil {
		return nil, err
	}

	// Find and iterate over each table row
	doc.Find("tr").Each(func(index int, item *goquery.Selection) {
		// Skip the header row
		if index > 0 {
			subCategory := item.Find("td:nth-child(1)").Text()
			detailData := item.Find("td:nth-child(2)").Text()

			data = append(data, SubCategory{
				SubCategoryName: subCategory,
				Details:         detailData,
			})
		}
	})

	return data, nil
}
