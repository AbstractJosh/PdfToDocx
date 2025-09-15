using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

// --- Spire aliases to avoid type/enum collisions ---
using PdfDocument = Spire.Pdf.PdfDocument;
using DocDocument = Spire.Doc.Document;

// Both libraries define FileFormat; alias them:
using PdfFileFormat = Spire.Pdf.FileFormat;
using DocFileFormat = Spire.Doc.FileFormat;

// ImportFormatMode lives under Spire.Doc.Documents
using Spire.Doc.Documents;

namespace WpfChunkPdfToDocx
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private async void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "Select a PDF"
            };

            if (ofd.ShowDialog() != true) return;

            BtnSelect.IsEnabled = false;
            Progress.Value = 0;
            TxtLog.Clear();
            TxtStatus.Text = "Status: reading PDF…";

            try
            {
                var inputPdfPath = ofd.FileName;
                var outputDocxPath = Path.Combine(
                    Path.GetDirectoryName(inputPdfPath)!,
                    Path.GetFileNameWithoutExtension(inputPdfPath) + "_converted.docx");

                // Run conversion on background thread
                await Task.Run(() => ConvertPdfToDocxInChunks(inputPdfPath, outputDocxPath, 10));

                TxtStatus.Text = $"Status: done → {outputDocxPath}";
                Log($"✅ Finished: {outputDocxPath}");
                MessageBox.Show($"Finished!\n\n{outputDocxPath}", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                TxtStatus.Text = "Status: error";
                Log("❌ " + ex);
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnSelect.IsEnabled = true;
                Progress.Value = 100;
            }
        }

        /// <summary>
        /// Splits the PDF into N-page chunks, converts each to DOCX, then merges to a single DOCX.
        /// </summary>
        private void ConvertPdfToDocxInChunks(string pdfPath, string finalDocxPath, int chunkSize)
        {
            // Ensure temp work folder
            string workDir = Path.Combine(Path.GetTempPath(), "PdfToDocxChunks_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(workDir);

            List<string> tempDocxParts = new();

            try
            {
                using (var probe = new PdfDocument())
                {
                    probe.LoadFromFile(pdfPath);
                    int totalPages = probe.Pages.Count;
                    int totalChunks = (int)Math.Ceiling(totalPages / (double)chunkSize);

                    DispatcherInvoke(() =>
                    {
                        Progress.Value = 0;
                        TxtStatus.Text = $"Status: converting {totalPages} pages in {totalChunks} chunks…";
                    });

                    for (int chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++)
                    {
                        int startPage = chunkIndex * chunkSize;             // 0-based
                        int endPageExclusive = Math.Min(startPage + chunkSize, totalPages);

                        string chunkPdfPath = Path.Combine(workDir, $"chunk_{chunkIndex + 1}.pdf");
                        string chunkDocxPath = Path.Combine(workDir, $"chunk_{chunkIndex + 1}.docx");

                        // Build a temporary PDF that only contains [startPage, endPageExclusive)
                        BuildChunkPdf(pdfPath, chunkPdfPath, startPage, endPageExclusive);

                        // Convert this chunk PDF → DOCX (Spire.PDF can save to DOCX directly)
                        using (var chunkDoc = new PdfDocument())
                        {
                            chunkDoc.LoadFromFile(chunkPdfPath);
                            chunkDoc.SaveToFile(chunkDocxPath, PdfFileFormat.DOCX); // alias avoids clash
                        }

                        tempDocxParts.Add(chunkDocxPath);

                        double progress = (chunkIndex + 1) * 100.0 / totalChunks;
                        DispatcherInvoke(() =>
                        {
                            TxtStatus.Text = $"Status: converted chunk {chunkIndex + 1}/{totalChunks} " +
                                             $"(pages {startPage + 1}-{endPageExclusive})";
                            Progress.Value = progress;
                            Log($"Chunk {chunkIndex + 1}: pages {startPage + 1}-{endPageExclusive} → {Path.GetFileName(chunkDocxPath)}");
                        });
                    }
                }

                // Merge all chunk DOCX into one final DOCX
                MergeDocxFiles(tempDocxParts, finalDocxPath);

                DispatcherInvoke(() =>
                {
                    TxtStatus.Text = "Status: merge complete";
                    Progress.Value = 100;
                    Log($"Merged {tempDocxParts.Count} parts → {finalDocxPath}");
                });
            }
            finally
            {
                // Cleanup temporary folder
                try { Directory.Delete(workDir, true); } catch { /* ignore */ }
            }
        }

        /// <summary>
        /// Creates a temporary PDF containing only pages [start, endExclusive).
        /// Strategy: reload original PDF, remove pages outside range.
        /// </summary>
        private static void BuildChunkPdf(string sourcePdfPath, string outPdfPath, int start, int endExclusive)
        {
            using var doc = new PdfDocument();
            doc.LoadFromFile(sourcePdfPath);

            // Remove pages we do NOT need (iterate from end for safety)
            for (int i = doc.Pages.Count - 1; i >= 0; i--)
            {
                if (i < start || i >= endExclusive)
                {
                    doc.Pages.RemoveAt(i);
                }
            }

            doc.SaveToFile(outPdfPath); // keep as PDF for the intermediary
        }

        /// <summary>
        /// Merges DOCX files back-to-back using Spire.Doc.
        /// </summary>
        private static void MergeDocxFiles(List<string> docxPaths, string outputDocxPath)
        {
            if (docxPaths == null || docxPaths.Count == 0)
                throw new InvalidOperationException("No DOCX parts to merge.");

            DocDocument? merged = null;

            foreach (var path in docxPaths)
            {
                var part = new DocDocument();
                part.LoadFromFile(path);

                if (merged == null)
                {
                    merged = part; // take ownership of first
                }
                else
                {
                    merged.AppendDocument(part, ImportFormatMode.KeepSourceFormatting);
                    part.Close();
                }
            }

            merged!.SaveToFile(outputDocxPath, DocFileFormat.Docx);
            merged.Close();
        }

        private void DispatcherInvoke(Action action)
        {
            Dispatcher.Invoke(action);
        }

        private void Log(string message)
        {
            TxtLog.AppendText($"{DateTime.Now:HH:mm:ss}  {message}{Environment.NewLine}");
            TxtLog.ScrollToEnd();
        }
    }
}
