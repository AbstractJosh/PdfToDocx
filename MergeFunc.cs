private static void MergeDocxFiles(List<string> docxPaths, string outputDocxPath)
{
    if (docxPaths == null || docxPaths.Count == 0)
        throw new InvalidOperationException("No DOCX parts to merge.");

    // Load first DOCX as base
    var merged = new DocDocument();
    merged.LoadFromFile(docxPaths[0]);

    for (int i = 1; i < docxPaths.Count; i++)
    {
        var part = new DocDocument();
        part.LoadFromFile(docxPaths[i]);

        // Append all child objects (paragraphs, tables, etc.)
        foreach (var obj in part.Sections[0].Body.ChildObjects)
        {
            merged.Sections[0].Body.ChildObjects.Add(obj.Clone());
        }

        part.Close();
    }

    merged.SaveToFile(outputDocxPath, DocFileFormat.Docx);
    merged.Close();
}
