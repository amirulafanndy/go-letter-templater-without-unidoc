package main

import (
	"fmt"

	"github.com/nguyenthenguyen/docx"
)

type Person struct {
	Name  string
	Date  string
	Email string
}

func main() {
	// Example data
	person := Person{
		Name:  "John Doe",
		Date:  "2024-10-01",
		Email: "johndoe@example.com",
	}

	// Read from docx file
	r, err := docx.ReadDocxFile("./template.docx")
	if err != nil {
		panic(err)
	}
	docxEditable := r.Editable()

	// Replace placeholders in the document
	docxEditable.Replace("{{name}}", person.Name, -1)
	docxEditable.Replace("{{date}}", person.Date, -1)
	docxEditable.Replace("{{email}}", person.Email, -1)

	// Optional: Replace links, headers, and footers
	docxEditable.ReplaceLink("http://example.com/", "https://github.com/nguyenthenguyen/docx", 1)
	docxEditable.ReplaceHeader("out with the old", "in with the new")
	docxEditable.ReplaceFooter("Change This Footer", "new footer")

	// Write the modified document to a new file
	docxEditable.WriteToFile("./new_result.docx")

	// Optional: Replace images
	// docxEditable.ReplaceImage("word/media/image1.png", "./new.png")
	// Uncomment the following line to replace the last image
	// imageIndex := docxEditable.ImagesLen()
	// docxEditable.ReplaceImage("word/media/image"+strconv.Itoa(imageIndex)+".png", "./new.png")

	// Close the document
	r.Close()

	fmt.Println("Document updated successfully and saved as new_result.docx")
}
