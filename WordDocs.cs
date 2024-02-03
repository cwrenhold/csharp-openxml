using DocumentFormat.OpenXml;
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
        }
    }

    public static void ReplaceTextInWordDocument(string outputDirectory)
    {
        var outputFilePath = Path.Combine(outputDirectory, "UpdatedFile.docx");
        File.Copy("InputFile.docx", outputFilePath, true);
        using (var wordDocument = WordprocessingDocument.Open(outputFilePath, true))
        {
            var mainPart = wordDocument.MainDocumentPart;
            var body = mainPart?.Document.Body;

            if (body == null || mainPart == null)
            {
                return;
            }

            var replacements = new Dictionary<string, string>
            {
                { "{pupil fullname}", "John Doe" },
                { "{teacher fullname}", "Jane Doe"},
                { "{summative result reading}", "Just At" },
                { "{summative result writing}", "Securely At" },
                { "{summative result mathematics}", "Below" }
            };

            // Replace the placeholders
            foreach (var text in body.Descendants<Text>())
            {
                foreach (var replacement in replacements)
                {
                    text.Text = text.Text.Replace(replacement.Key, replacement.Value);
                }

                if (text.Text.Contains("{") || text.Text.Contains("}"))
                {
                    System.Diagnostics.Debug.WriteLine($"Unreplaced placeholder: {text.Text}");
                }
            }


            // Save the document to the output directory
            wordDocument.Save();
        }
    }
}
