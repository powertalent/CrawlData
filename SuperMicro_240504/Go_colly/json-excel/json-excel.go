package main

import (
	"encoding/json"
	"fmt"
	"log"

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

func main() {
	jsonData := `
    [
        {
            "MainCategory": "Electronics",
            "SubCategories": [
                {
                    "SubCategory": "Laptops",
                    "DetailValue": "Dell: $1200"
                },
                {
                    "SubCategory": "Cameras",
                    "DetailValue": "Canon: $500"
                }
            ]
        },
        {
            "MainCategory": "Appliances",
            "SubCategories": [
                {
                    "SubCategory": "Refrigerators",
                    "DetailValue": "Whirlpool: $800"
                },
                {
                    "SubCategory": "Microwaves",
                    "DetailValue": "Samsung: $300"
                }
            ]
        }
    ]
    `
	var categories []MainCategory
	err := json.Unmarshal([]byte(jsonData), &categories)
	if err != nil {
		log.Fatalf("Error parsing JSON: %v", err)
	}

	workbook := xlsx.NewFile()
	sheet, err := workbook.AddSheet("Data")
	if err != nil {
		log.Fatalf("Error adding sheet to XLSX file: %v", err)
	}

	// Process each main category
	for _, category := range categories {
		mcRow := sheet.AddRow()
		mcCell := mcRow.AddCell()
		mcCell.Value = category.CategoryName
		mcCell.Merge(1, 0) // Merge two cells for Main Category for clarity

		// SubCategories and Details
		for _, subCat := range category.SubCategories {
			scRow := sheet.AddRow()
			scCell := scRow.AddCell()
			scCell.Value = subCat.SubCategoryName
			detailCell := scRow.AddCell()
			detailCell.Value = subCat.DetailValue
		}
	}

	// Save the workbook
	err = workbook.Save("FormattedData.xlsx")
	if err != nil {
		log.Fatalf("Error saving XLSX file: %v", err)
	}

	fmt.Println("Excel file saved successfully.")
}
