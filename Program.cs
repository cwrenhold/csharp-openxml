using CsharpOpenXml;
using Path = System.IO.Path;

var outputDirectory = Path.Combine(Environment.CurrentDirectory, "output");

// Empty the output directory
if (Directory.Exists(outputDirectory))
{
    Directory.Delete(outputDirectory, true);
}

Directory.CreateDirectory(outputDirectory);

// Create a new Word document from scratch
WordDocs.CreateWordDocumentFromScratch(outputDirectory);

// Copy the input document to the output directory, then replace the placeholders
WordDocs.ReplaceTextInWordDocument(outputDirectory);

PowerPointDocs.CreatePowerPointDocumentFromScratch(outputDirectory);

PowerPointDocs.ReplaceTextInPowerPointPresentation(outputDirectory);
