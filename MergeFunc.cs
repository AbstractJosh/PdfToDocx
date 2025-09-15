// at the top of the file you already have:
// using Spire.Doc.Documents;                  // needed for DocumentObject, BreakType
// using DocDocument = Spire.Doc.Document;
// using DocFileFormat = Spire.Doc.FileFormat;

private static void MergeDocxFiles(List<string> docxPaths, string outputDocxPath)
{
    if (docxPaths == null || docxPaths.Count == 0)
        throw new InvalidOperationException("No DOCX parts to merge.");

    var merged = new DocDocument();
    merged.LoadFromFile(docxPaths[0]);

    // Append each subsequent document's content
    for (int i = 1; i < docxPaths.Count; i++)
    {
        var part = new DocDocument();
        part.LoadFromFile(docxPaths[i]);

        // copy every section's body objects (paragraphs, tables, pictures, etc.)
        for (int s = 0; s < part.Sections.Count; s++)
        {
            var body = part.Sections[s].Body;
            for (int j = 0; j < body.ChildObjects.Count; j++)
            {
                DocumentObject cloned = body.ChildObjects[j].Clone();
                merged.Sections[0].Body.ChildObjects.Add(cloned);
            }

            // optional: page break between sections/chunks
            var pb = merged.Sections[0].AddParagraph();
            pb.AppendBreak(BreakType.PageBreak);
        }

        part.Close();
    }

    merged.SaveToFile(outputDocxPath, DocFileFormat.Docx);
    merged.Close();
}
