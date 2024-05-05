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

	workbook := xlsx.NewFile()
	sheet, err := workbook.AddSheet("Data")
	if err != nil {
		log.Fatalf("Error adding sheet to XLSX file: %v", err)
	}

	header := sheet.AddRow()
	header.AddCell() // Add an empty cell for the category column
	for _, product := range products {
		cell := header.AddCell()
		cell.Value = product.ProductName
	}

	// Map to store unique subcategories for each main category across all products
	mainCategorySubCats := make(map[string]map[string]bool)

	// First collect all unique subcategories for each main category across all products
	for _, product := range products {
		for _, category := range product.MainCategories {
			if _, ok := mainCategorySubCats[category.CategoryName]; !ok {
				mainCategorySubCats[category.CategoryName] = make(map[string]bool)
			}
			for _, subCategory := range category.SubCategories {
				mainCategorySubCats[category.CategoryName][subCategory.SubCategoryName] = true
			}
		}
	}

	// Iterate through each main category and list each unique subcategory once
	for mcName, subCats := range mainCategorySubCats {
		mcRow := sheet.AddRow()
		mcCell := mcRow.AddCell()
		mcCell.Value = mcName
		mcCell.Merge(len(products), 0) // Merge across product columns for the main category label

		subCategoryList := make([]string, 0, len(subCats))
		for subCatName := range subCats {
			subCategoryList = append(subCategoryList, subCatName)
		}
		sort.Strings(subCategoryList)

		for _, subCatName := range subCategoryList {
			subCatRow := sheet.AddRow()
			subCatRow.AddCell().Value = subCatName

			// Populate detail values across all products for this subcategory
			for _, product := range products {
				found := false
				for _, category := range product.MainCategories {
					if category.CategoryName == mcName {
						for _, subCat := range category.SubCategories {
							if subCat.SubCategoryName == subCatName {
								subCatRow.AddCell().Value = subCat.DetailValue
								found = true
								break
							}
						}
					}
					if found {
						break
					}
				}
				if !found {
					subCatRow.AddCell().Value = "" // Empty cell if this product's main category doesn't have the subcategory
				}
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
