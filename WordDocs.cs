using System.Text.RegularExpressions;
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
                { "#{school name}#", "Replacement Academy" },
                { "#{pupil first name}#", "John" },
                { "#{pupil last name}#", "Doe" },
                { "#{year group}#", "Year 6" },
                { "#{class name}#", "Dolphins" },
                { "#{academic year}#", "2023-24"},

                { "#{atl grade label}#", "Effort" },
                { "#{atl grade 1 grade}#", "1 - Gooderer" },
                { "#{atl grade 1 descriptor}#", "Good - School defined description for this" },
                { "#{atl grade 2 grade}#", "2 - Good" },
                { "#{atl grade 2 descriptor}#", "Good - School defined description for this" },
                { "#{atl grade 3 grade}#", "3 - \"Acceptable\"" },
                { "#{atl grade 3 descriptor}#", "Good - School defined description for this" },

                { "#{reading summative result reading}#", "Just At" },
                { "#{summative result writing}#", "Securely At" },
                { "#{summative result mathematics}#", "Below" },
            };

        foreach (var item in subjectResults)
        {
            var subjectName = item.Label.ToLower();

            // For each subject, create the replacements in the format of:
            // #{subject reading label}#
            // #{subject reading result}#
            // #{subject reading other grade}#
            // #{subject reading target 1 text}#
            // #{subject reading target 1 colour}#
            replacements.Add($"#{{subject {subjectName} label}}#", item.Label);
            replacements.Add($"#{{subject {subjectName} result}}#", item.Result);
            replacements.Add($"#{{subject {subjectName} other grade}}#", item.OtherGrade);

            // we will get the 3 targets but need to replace the placeholders for all
            for (var j = 0; j < 3; j++)
            {
                var target = item.Targets.ElementAtOrDefault(j);
                replacements.Add($"#{{subject {subjectName} target {j + 1} text}}#", target.Text ?? "");
                replacements.Add($"#{{subject {subjectName} target {j + 1} colour}}#", target.Colour ?? "");
            }
        }

        var outputFilePath = Path.Combine(outputDirectory, "UpdatedFile.docx");
        File.Copy("ParentReportTemplate_Four.docx", outputFilePath, true);
        using (var wordDocument = WordprocessingDocument.Open(outputFilePath, true))
        {
            ReplaceStrings(wordDocument.MainDocumentPart, wordDocument.MainDocumentPart?.Document?.Body, replacements);

            foreach (var headerPart in wordDocument.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
            {
                ReplaceStrings(wordDocument.MainDocumentPart, headerPart.Header, replacements);
            }

            foreach (var footerPart in wordDocument.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
            {
                ReplaceStrings(wordDocument.MainDocumentPart, footerPart.Footer, replacements);
            }

            // Save the document to the output directory
            wordDocument.Save();
        }
    }

    public static void ReplaceStrings(OpenXmlPart? part, OpenXmlCompositeElement? partElement, Dictionary<string, string> replacements)
    {
        if (part is null || partElement is null)
        {
            return;
        }

        Dictionary<string, List<Text>> textReplacementNodes = new();

        // Replace the placeholders
        foreach (var text in partElement.Descendants<Text>())
        {
            foreach (var replacement in replacements)
            {
                text.Text = text.Text.Replace(replacement.Key, replacement.Value);
            }

            if (text.Text.Contains("#{") || text.Text.Contains("}#"))
            {
                var parent = text.Parent;

                var maxDepth = 10;
                while (parent is not null
                    && !replacements.Any(r => parent.InnerText.Contains(r.Key))
                    && maxDepth > 0)
                {
                    parent = parent?.Parent;
                    maxDepth--;
                }

                if (parent is not null)
                { 
                    System.Diagnostics.Debug.WriteLine($"Unreplaced text in node: {parent.InnerText}");

                    foreach (var replacement in replacements.Where(r => parent.InnerText.Contains(r.Key)))
                    {
                        // Now that we know which parent contains the unreplaced text, we can replace it
                        var allElements = parent.Elements().ToList();
                        var textNodes = parent.Descendants<Text>().ToList();

                        // Find the nodes which make up the text
                        var nodes = FindStartAndEndNodes(parent.Descendants<Text>().ToList(), replacement.Key);

                        if (nodes is null)
                        {
                            continue;
                        }

                        var (startNode, endNode) = nodes.Value;

                        // Find the index of the elements which contain the start and end nodes
                        var startIndex = allElements.IndexOf(allElements.First(c => c.Descendants().Contains(startNode)));
                        var endIndex = allElements.IndexOf(allElements.First(c => c.Descendants().Contains(endNode)));

                        // Find the inner text of all nodes between the start and end nodes
                        var innerText = string.Join("", allElements.Skip(startIndex).Take(endIndex - startIndex + 1).Select(c => c.InnerText));

                        innerText = innerText.Replace(replacement.Key, replacement.Value);

                        startNode.Text = innerText;

                        // Remove the nodes between the start and end nodes, starting from the end node
                        for (var i = endIndex; i > startIndex; i--)
                        {
                            parent.RemoveChild(allElements[i]);
                        }

                        if (textReplacementNodes.ContainsKey(replacement.Key))
                        {
                            textReplacementNodes[replacement.Key].Add(startNode);
                        }
                        else
                        {
                            textReplacementNodes.Add(replacement.Key, new List<Text> { startNode });
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"Replaced text in node: {parent.InnerText}");
                }
            }
        }
    }

    private static (Text startNode, Text endNode)? FindStartAndEndNodes(List<Text> textNodes, string key, int? nodeTextOffset = null)
    {
        // Find the index of the first node which contains part of the key
        var startIndex = -1;
        var currentTextIndex = -1;

        for (var i = 0; i < textNodes.Count; i++)
        {
            var text = i == 0 && nodeTextOffset is not null
                ? textNodes[i].Text.Substring(nodeTextOffset.Value)
                : textNodes[i].Text;

            var textIndex = text.IndexOf(key[0]);
            if (textIndex >= 0)
            {
                startIndex = i;
                currentTextIndex = textIndex;
                break;
            }
        }

        // If the start index is -1, the key was not found
        if (startIndex == -1 || currentTextIndex == -1)
        {
            return null;
        }

        System.Diagnostics.Debug.WriteLine($"Processing key: {key} at start index: {startIndex} and text index: {currentTextIndex}");

        var currentNodeIndex = startIndex;
        var tokenLeftToFind = key;
        var currentNodeText = textNodes[currentNodeIndex].Text.Substring(currentTextIndex);

        // Remove the text from the token as we find it
        while (tokenLeftToFind.Length > 0)
        {
            if (currentNodeText.Length == 0)
            {
                currentNodeIndex++;

                // If we've reached the end of the nodes, the token is no longer valid
                if (currentNodeIndex >= textNodes.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"Token no longer valid, reached the end of the nodes.");
                    return null;
                }

                currentNodeText = textNodes[currentNodeIndex].Text;
            }

            var nextCharacter = currentNodeText[0];
            var nextTokenCharacter = tokenLeftToFind[0];

            // If the next index is one after the current index, we're still in the same node and the token is still valid
            if (nextCharacter == nextTokenCharacter)
            {
                currentNodeText = currentNodeText.Substring(1);
                tokenLeftToFind = tokenLeftToFind.Substring(1);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Token no longer valid, restarting from the next node, at index {currentNodeIndex}.");
                // If the next index is not one after the current index, the token is no longer valid, so we need to reset and continue from where we are
                var newNodeSet = textNodes.Skip(currentNodeIndex).ToList();

                return FindStartAndEndNodes(newNodeSet, key, currentTextIndex + 1);
            }
        }

        var startNode = textNodes[startIndex];
        var endNode = textNodes[currentNodeIndex];

        return (startNode, endNode);
    }
}
