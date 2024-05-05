package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"sort"

	"github.com/tealeg/xlsx/v3"
)

type SubCategory struct {
	SubCategoryName string `json:"SubCategory"`
	DetailValue     string `json:"DetailValue"`
}

type MainCategory struct {
	CategoryName  string        `json:"MainCategory"`
	SubCategories []SubCategory `json:"SubCategories"`
}

type Product struct {
	ProductName    string         `json:"ProductName"`
	MainCategories []MainCategory `json:"MainCategories"`
}

func main() {
	jsonData, err := ioutil.ReadFile("data.json")
	if err != nil {
		log.Fatalf("Error opening JSON file: %v", err)
	}

	var products []Product
	err = json.Unmarshal(jsonData, &products)
	if err != nil {
		log.Fatalf("Error parsing JSON: %v", err)
	}

	// Collect all unique subcategories across all products
	uniqueSubCategories := make(map[string]struct{})
	for _, product := range products {
		for _, category := range product.MainCategories {
			for _, subCategory := range category.SubCategories {
				uniqueSubCategories[subCategory.SubCategoryName] = struct{}{}
			}
		}
	}

	// Convert map keys to slice for sorting and indexed access
	subCategoryList := make([]string, 0, len(uniqueSubCategories))
	for subCat := range uniqueSubCategories {
		subCategoryList = append(subCategoryList, subCat)
	}
	sort.Strings(subCategoryList) // Sort subcategories alphabetically

	// Create a new Excel file
	workbook := xlsx.NewFile()
	sheet, err := workbook.AddSheet("Data")
	if err != nil {
		log.Fatalf("Error adding sheet to XLSX file: %v", err)
	}

	// Create a header row for product names starting from the second column
	header := sheet.AddRow()
	header.AddCell() // Add an empty cell for the subcategory column
	for _, product := range products {
		cell := header.AddCell()
		cell.Value = product.ProductName
	}

	// Create rows for each subcategory
	for _, subCatName := range subCategoryList {
		row := sheet.AddRow()
		row.AddCell().Value = subCatName // Subcategory name in the first column

		// Fill in the detail values for each product
		for _, product := range products {
			productHasSubCat := false
			for _, category := range product.MainCategories {
				for _, subCat := range category.SubCategories {
					if subCat.SubCategoryName == subCatName {
						row.AddCell().Value = subCat.DetailValue
						productHasSubCat = true
						break
					}
				}
				if productHasSubCat {
					break
				}
			}
			if !productHasSubCat {
				row.AddCell().Value = "" // Empty cell if the product doesn't have this subcategory
			}
		}
	}

	// Save the workbook
	err = workbook.Save("FormattedData.xlsx")
	if err != nil {
		log.Fatalf("Error saving XLSX file: %v", err)
	}

	fmt.Println("Excel file saved successfully.")
}
