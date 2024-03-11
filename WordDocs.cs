using System.Text.RegularExpressions;
using Codeuctivity.OpenXmlPowerTools;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CsharpOpenXml;

public static class WordDocs
{


    public static void CreateWordDocumentFromScratch(string outputDirectory)
    {
        using (var wordDocument = WordprocessingDocument.Create(Path.Combine(outputDirectory, "HelloWorld.docx"), WordprocessingDocumentType.Document))
        {
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Append(body);
            body.Append(new Paragraph(new Run(new Text("Hello, World!"))));


            // insert white space
            body.Append(new Paragraph(new Run(new Text(""))));

            // Make the table width 100% of the page width.
            var tableWidth = new TableWidth() { Width = "100", Type = TableWidthUnitValues.Pct };
            var tableProp = new TableProperties();
            var tableStyle = new TableStyle() { Val = "TableGrid" };
            tableProp.Append(tableStyle, tableWidth);

            // list of subjects with results and effort
            var resultList = new List<(string Subject, string Result, string Effort)>
            {
                ("Reading", "Just At", "Good"),
                ("Writing", "Securely At", "Good"),
                ("Mathematics", "Below", "Good")
            };

            // // Table header rpw
            // var subjectTable = new Table(new TableRow(
            //      new TableCell(GenerateTableCellPropsWithWidth("60"), new Paragraph(new Run(new Text("Subject")))),
            //      new TableCell(GenerateTableCellPropsWithWidth("20"), new Paragraph(new Run(new Text("Result")))),
            //      new TableCell(GenerateTableCellPropsWithWidth("20"), new Paragraph(new Run(new Text("Effort"))))
            //  ));

            // // add the rest of the rows
            // subjectTable.Append(resultList.Select(x => new TableRow(
            //     new TableCell(new Paragraph(new Run(new Text(x.Subject)))),
            //     new TableCell(new Paragraph(new Run(new Text(x.Result)))),
            //     new TableCell(new Paragraph(new Run(new Text(x.Effort))))
            // )));

            // body.Append(subjectTable);

            // // insert white space
            // body.Append(new Paragraph(new Run(new Text(""))));

            // // New test table
            // var tbl = new Table();





            // tbl.AppendChild(tableProp);

            // //Add n columns to table
            // var tg = new TableGrid(new GridColumn(), new GridColumn());

            // tbl.AppendChild(tg);

            // var tr1 = new TableRow();

            // //I Manually adjust width of the first column
            // var tc1 = new TableCell(GenerateTableCellPropsWithWidth("270"), new Paragraph(new Run(new Text("№"))));

            // //All other column are adjusted based on their content
            // var tc2 = new TableCell(GenerateTableCellPropsWithWidth(), new Paragraph(new Run(new Text("Title"))));

            // tr1.Append(tc1, tc2);
            // tbl.AppendChild(tr1);

            // //This method is only used for headers, while regular rows cells contain no TableCellProperties
            // TableCellProperties GenerateTableCellPropsWithWidth(string width = "")
            // {
            //     // if width is null, the TableCellWidth will be set to Auto
            //     var tableCell = string.IsNullOrEmpty(width)
            //         ? new TableCellWidth { Type = TableWidthUnitValues.Auto }
            //         : new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width };

            //     TableCellProperties tcp = new TableCellProperties();
            //     tcp.AppendChild(tableCell);
            //     return tcp;
            // }

            // body.AppendChild(tbl);

        }
    }

    public static void ReplaceTextInWordDocument(string outputDirectory)
    {
        // Create a dictionary of replacements
        /*
            Reading
            Writing
            Mathematics
            Science
            Spoken Language
            Art and Design
            Computing
            Design and Technology
            History
            Geography
            Languages
            Music
            PSHE
            Physical Education
            Religious Education
            */

        // create  list of subjects
        var subjectResults = new List<(string Label, string Result, string OtherGrade, (string Text, string Colour)[] Targets)>();
        subjectResults.AddRange(
            new[]
            {
                    ("Reading", "Just At", "Good", new[] { ("Can read the word \"the\"", "Red"), ("Understands what a book is", "Red"), ("Understands the letter P", "Yellow") }),
                    ("Writing", "Securely At", "Good", new[] { ("Can pick up a pen", "Red"), ("Can write own name", "Red"), ("Understands the letter P", "Yellow") }),
                    ("Mathematics", "Below", "Good", new[] { ("Can count to 1", "Red"), ("Understands decimal", "Red")}),
                    ("Science", "Just At", "Good", new[] { ("Understands science", "Red") }),
                    ("Spoken Language", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Art and Design", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Computing", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Design and Technology", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("History", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Geography", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Languages", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Music", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("PSHE", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Physical Education", "Just At", "Good", Array.Empty<(string, string)>()),
                    ("Religious Education", "Just At", "Good", Array.Empty<(string, string)>()),
            }
        );





        var replacements = new Dictionary<string, string>
            {
                { "{school name}", "Replacement Academy" },
                { "{pupil first name}", "John" },
                { "{pupil last name}", "Doe" },
                { "{year group}", "Year 6" },
                { "{class name}", "Dolphins" },
                { "{academic year}", "2023-24"},

                { "{other_grade_label}", "Effort" },
                { "{other_grade_1_grade}", "1 - Gooderer" },
                { "{other grade 1 descriptor}", "Good - School defined description for this" },
                { "{other grade 2 grade}", "2 - Good" },
                { "{other grade 2 descriptor}", "Good - School defined description for this" },
                { "{other grade 3 grade}", "3 - \"Acceptable\"" },
                { "{other grade 3 descriptor}", "Good - School defined description for this" },

                { "{reading summative result reading}", "Just At" },
                { "{summative result writing}", "Securely At" },
                { "{summative result mathematics}", "Below" },
            };

        foreach (var item in subjectResults)
        {
            replacements.Add($"{{subject {item.Label} label}}", item.Label);
            replacements.Add($"{{subject {item.Label} result}}", item.Result);
            replacements.Add($"{{subject {item.Label} other grade}}", item.OtherGrade);

            // we will get the 3 targets but need to replace the placeholders for all
            for (var j = 0; j < 3; j++)
            {
                var target = item.Targets.ElementAtOrDefault(j);
                replacements.Add($"{{subject {item.Label} target {j + 1} text}}", target.Text ?? "");
                replacements.Add($"{{subject {item.Label} target {j + 1} colour}}", target.Colour ?? "");
            }
        }

        var outputFilePath = Path.Combine(outputDirectory, "UpdatedFile.docx");
        File.Copy("ParentReportTemplate_Four.docx", outputFilePath, true);
        using (var wordDocument = WordprocessingDocument.Open(outputFilePath, true))
        {
            var mainPart = wordDocument.MainDocumentPart;
            var body = mainPart?.Document.Body;

            if (body == null || mainPart == null)
            {
                return;
            }

            foreach (var replacement in replacements)
            {
                TextReplacer.SearchAndReplace(wordDocument, replacement.Key, replacement.Value, false);
            }

            // // Replace the placeholders
            // foreach (var text in body.Descendants<Text>())
            // {
            //     foreach (var replacement in replacements)
            //     {
            //         text.Text = text.Text.Replace(replacement.Key, replacement.Value);

            //     }

            //     if (text.Text.Contains("{") || text.Text.Contains("}"))
            //     {
            //         System.Diagnostics.Debug.WriteLine($"Unreplaced placeholder: {text.Text}");
            //     }
            // }

            // Save the document to the output directory
            wordDocument.Save();
        }
    }
}
