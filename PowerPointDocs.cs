using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Drawing.Text;

namespace CsharpOpenXml;

public static class PowerPointDocs
{
    public static void CreatePowerPointDocumentFromScratch(string outputDirectory)
    {
        using (var presentationDocument = PresentationDocument.Create(System.IO.Path.Combine(outputDirectory, "HelloWorld.pptx"), PresentationDocumentType.Presentation))
        {
            var presentationPart = presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
            var slideIdList = presentationPart.Presentation.AppendChild(new SlideIdList());
            slideIdList.AppendChild(new SlideId { Id = 256, RelationshipId = presentationPart.GetIdOfPart(slidePart) });

            // Save the document to the output directory
            presentationDocument.Save();
        }
    }

    public static void ReplaceTextInPowerPointPresentation(string outputDirectory)
    {
        var outputFilePath = System.IO.Path.Combine(outputDirectory, "UpdatedFile.pptx");
        File.Copy("InputFile.pptx", outputFilePath, true);
        using (var presentationDocument = PresentationDocument.Open(outputFilePath, true))
        {
            var presentationPart = presentationDocument.PresentationPart;

            foreach (var slidePart in presentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>())
            {
                if (slidePart == null || presentationPart == null)
                {
                    return;
                }

                var replacements = new Dictionary<string, string>
                {
                    { "{pupil fullname}", "John Doe" },
                    { "{teacher fullname}", "Jane Doe"},
                    { "{summative result reading}", "Just At" },
                    { "{summative result writing}", "Securely At" },
                    { "{summative result mathematics}", "Below" },
                    { "{summative atorabove reading_writing_mathematics}", "65.3%"},
                    { "{school name}", "School Name" },
                    { "{registration name}", "Dolphins"},
                    { "{registration pupil_count}", "31"},
                };

                // Replace the placeholders
                foreach (var text in slidePart.Slide.Descendants<Text>())
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
            }

            // Save the document to the output directory
            presentationDocument.Save();
        }
    }
}
