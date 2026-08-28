using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using MSOfficeAIAssistant.Attachments;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI
{
    public partial class DocumentCompareWindow : Window
    {
        private string _fileA;
        private string _fileB;
        private string _diffText;

        public DocumentCompareWindow()
        {
            InitializeComponent();
        }

        public static void ShowFor(string fileA = null, string fileB = null)
        {
            var win = new DocumentCompareWindow();
            if (!string.IsNullOrWhiteSpace(fileA)) { win._fileA = fileA; win.TxtFileA.Text = Path.GetFileName(fileA); }
            if (!string.IsNullOrWhiteSpace(fileB)) { win._fileB = fileB; win.TxtFileB.Text = Path.GetFileName(fileB); }
            win.ShowDialog();
        }

        private void BtnBrowseA_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Document A",
                Filter = "Supported (*.docx;*.pdf;*.txt;*.csv;*.md)|*.docx;*.pdf;*.txt;*.csv;*.md;*.json;*.xml|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                _fileA = dlg.FileName;
                TxtFileA.Text = _fileA;
            }
        }

        private void BtnBrowseB_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Document B",
                Filter = "Supported (*.docx;*.pdf;*.txt;*.csv;*.md)|*.docx;*.pdf;*.txt;*.csv;*.md;*.json;*.xml|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                _fileB = dlg.FileName;
                TxtFileB.Text = _fileB;
            }
        }

        private async void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_fileA) || string.IsNullOrWhiteSpace(_fileB))
            {
                MessageBox.Show("Please select both documents.", "Document Comparison", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                TxtDiff.Text = "Extracting and comparing...";
                string textA = await ExtractText(_fileA);
                string textB = await ExtractText(_fileB);
                // DiffEngine's LCS is O(m*n) in lines; run it off the UI thread and bail out with a
                // clear message instead of hanging/OOM-ing on a line-dense extraction (DiffEngine.cs).
                List<DiffPiece> diff = await Task.Run(() => DiffEngine.DiffLinesOrNull(textA, textB));
                if (diff == null)
                {
                    TxtDiff.Text = string.Format(
                        "One or both documents have too many lines to compare (limit {0} lines per side). " +
                        "Try a shorter excerpt or a smaller document.", DiffEngine.MaxLinesPerSide);
                    return;
                }
                _diffText = DiffEngine.RenderPlain(diff);
                TxtDiff.Text = _diffText;
                if (string.IsNullOrWhiteSpace(_diffText)) TxtDiff.Text = "No differences (or empty files).";
            }
            catch (Exception ex)
            {
                Logger.Error("DocumentCompareWindow.Compare failed", ex);
                MessageBox.Show(string.Format("Comparison failed: {0}", ex.Message), "Document Comparison", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtDiff.Text = string.Format("Error: {0}", ex.Message);
            }
        }

        private async Task<string> ExtractText(string path)
        {
            var block = await AttachmentExtractor.ExtractAsync(path);
            return block.ExtractedText ?? string.Empty;
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_diffText))
            {
                MessageBox.Show("Nothing to copy - run a comparison first.", "Document Comparison", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            try
            {
                Clipboard.SetText(_diffText);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not copy: {0}", ex.Message), "Document Comparison", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
