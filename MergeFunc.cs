private static void MergeDocxFiles(List<string> docxPaths, string outputDocxPath)
{
    if (docxPaths == null || docxPaths.Count == 0)
        throw new InvalidOperationException("No DOCX parts to merge.");

    // Load first as the base
    var merged = new DocDocument();
    merged.LoadFromFile(docxPaths[0]);

    // Append each subsequent DOCX manually
    for (int i = 1; i < docxPaths.Count; i++)
    {
        var part = new DocDocument();
        part.LoadFromFile(docxPaths[i]);

        foreach (Section sec in part.Sections)
        {
            // Clone each section and add to merged doc
            merged.Sections.Add(sec.Clone());
        }

        part.Close();
    }

    merged.SaveToFile(outputDocxPath, DocFileFormat.Docx);
    merged.Close();
}
