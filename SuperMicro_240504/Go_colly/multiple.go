package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"
	"sync"

	"github.com/PuerkitoBio/goquery"
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

func scrapeURL(url string, wg *sync.WaitGroup, ch chan<- []string) {
	defer wg.Done()

	// Simulate HTML content fetch
	// Note: Replace with real fetching logic as necessary
	htmlContent := `<table>
        <tr><th>Header 1</th><th>Header 2</th></tr>
        <tr><td>Row 1 Col 1</td><td><a href="link1.html">Link 1</a></td></tr>
        <tr><td>Row 2 Col 1</td><td><a href="link2.html">Link 2</a></td></tr>
    </table>`

	doc, err := goquery.NewDocumentFromReader(strings.NewReader(htmlContent))
	if err != nil {
		fmt.Println("Error loading HTML: ", err)
		return
	}

	var links []string
	links = append(links, url)
	doc.Find("a").Each(func(i int, s *goquery.Selection) {
		if i < 10 { // Limit to first 10 links
			link, _ := s.Attr("href")
			links = append(links, link)
		}
	})

	ch <- links
}

func main() {
	// Read URLs from file
	urls, err := readURLsFromFile("link.txt")
	if err != nil {
		fmt.Println("Error reading URLs:", err)
		return
	}

	var wg sync.WaitGroup
	linkChannel := make(chan []string, len(urls))

	for _, url := range urls {
		wg.Add(1)
		go scrapeURL(url, &wg, linkChannel)
	}

	wg.Wait()
	close(linkChannel)

	// Create a new Excel file
	f := excelize.NewFile()
	sheetName := "Links"
	index, err := f.NewSheet(sheetName)
	f.SetActiveSheet(index)

	// Read from the channel and write to Excel
	row := 1
	for links := range linkChannel {
		for _, link := range links {
			cellAddress, _ := excelize.CoordinatesToCellName(1, row)
			f.SetCellValue(sheetName, cellAddress, link)
			row++
		}
	}

	// Save Excel file
	if err := f.SaveAs("Links.xlsx"); err != nil {
		fmt.Println("Error saving file: ", err)
	} else {
		fmt.Println("Excel file created successfully.")
	}
}

// func main() {

// 	// Create a new Excel file
// 	f := excelize.NewFile()
// 	defer func() {
// 		if err := f.Close(); err != nil {
// 			fmt.Println(err)
// 		}
// 	}()
// 	// Create a new sheet
// 	index, err := f.NewSheet("Sheet1")
// 	if err != nil {
// 		fmt.Println(err)
// 		return
// 	}
// 	// Set value of a cell
// 	f.SetCellValue("Sheet1", "A1", "Link Text")
// 	f.SetCellValue("Sheet1", "B1", "URL")

// 	// Counter for Excel rows
// 	row := 2

// 	// Instantiate default collector
// 	c := colly.NewCollector(
// 		colly.AllowedDomains("www.supermicro.com"),
// 	)

// 	// On every a element which has href attribute call callback
// 	c.OnHTML(".scrollable-table [class^=specHeader]", func(e *colly.HTMLElement) {
// 		header := e.Text
// 		// Print link
// 		fmt.Printf("Link found: %q -> %s\n", e.Text, header)

// 		// Write data to Excel
// 		cellA := fmt.Sprintf("A%d", row)
// 		cellB := fmt.Sprintf("B%d", row)
// 		f.SetCellValue("Sheet1", cellA, e.Text)
// 		f.SetCellValue("Sheet1", cellB, header)
// 		// parentHtml, err := e.DOM.Parent().Parent().Parent().Parent().Html() // Capture both the HTML and the error
// 		parentHtml, err := e.DOM.Closest("table").Html()
// 		if err != nil {
// 			fmt.Println("Error getting parent HTML:", err)
// 			return
// 		}
// 		f.SetCellValue("Sheet1", "C1", parentHtml)
// 		// Visit link found on page
// 		// Only those links are visited which are in AllowedDomains
// 		// c.Visit(e.Request.AbsoluteURL(link))
// 	})

// 	// Before making a request print "Visiting ..."
// 	c.OnRequest(func(r *colly.Request) {
// 		fmt.Println("Visiting", r.URL.String())
// 	})

// 	c.Visit("https://www.supermicro.com/ja/products/motherboard/A1SAM-2750F")

// 	// Set active sheet of the workbook
// 	f.SetActiveSheet(index)
// 	// Save spreadsheet by the given path.
// 	if err := f.SaveAs("Book1.xlsx"); err != nil {
// 		fmt.Println(err)
// 	}
// }
