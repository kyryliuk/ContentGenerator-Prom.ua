using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using GemBox.Document;
using Microsoft.Win32;

namespace Content
{
    public partial class EditWindow : Window
    {
        
        public EditWindow()
        {
            InitializeComponent();

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }
        public static string infocontent;
        public EditWindow(string content)
        {
            InitializeComponent();
            richTextBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(content)));
            infocontent = content;
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }
        private void Open(object sender, ExecutedRoutedEventArgs e)
        {
            var dialog = new OpenFileDialog()
            {
                AddExtension = true,
                Filter =
                    "All Documents (*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.htm;*.html;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.txt|" +
                    "Word Documents (*.docx)|*.docx|" +
                    "Word Macro-Enabled Documents (*.docm)|*.docm|" +
                    "Word 97-2003 Documents (*.doc)|*.doc|" +
                    "Word Templates (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Templates (*.dotm)|*.dotm|" +
                    "Word 97-2003 Templates (*.dot)|*.dot|" +
                    "Web Pages (*.htm;*.html)|*.htm;*.html|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Text Files (*.txt)|*.txt"
            };

            if (dialog.ShowDialog() == true)
                using (var stream = new MemoryStream())
                {
                    // Convert input file to RTF stream.
                    DocumentModel.Load(dialog.FileName).Save(stream, SaveOptions.RtfDefault);

                    stream.Position = 0;
                    
                    // Load RTF stream into RichTextBox.
                    var textRange = new TextRange(this.richTextBox.Document.ContentStart, this.richTextBox.Document.ContentEnd);
                    textRange.Load(stream, DataFormats.Rtf);
                }
        }

        protected override void OnClosed(EventArgs e)
        {
            
            base.OnClosed(e);
        }
        
        public void Save(object sender, ExecutedRoutedEventArgs e)
        {   
                infocontent = "";
                var textRange = new TextRange(this.richTextBox.Document.ContentStart, this.richTextBox.Document.ContentEnd);
                infocontent = textRange.Text;
           // MessageBox.Show(infocontent);
        }

        private void Cut(object sender, ExecutedRoutedEventArgs e)
        {
            this.Copy(sender, e);

            // Clear selection.
            this.richTextBox.Selection.Text = string.Empty;
        }

        private void Copy(object sender, ExecutedRoutedEventArgs e)
        {
            using (var stream = new MemoryStream())
            {
                // Save RichTextBox selection to RTF stream.
                this.richTextBox.Selection.Save(stream, DataFormats.Rtf);

                stream.Position = 0;

                // Save RTF stream to clipboard.
                DocumentModel.Load(stream, LoadOptions.RtfDefault).Content.SaveToClipboard();
            }
        }

        private void Paste(object sender, ExecutedRoutedEventArgs e)
        {
            using (var stream = new MemoryStream())
            {
                // Save RichTextBox content to RTF stream.
                var textRange = new TextRange(this.richTextBox.Document.ContentStart, this.richTextBox.Document.ContentEnd);
                textRange.Save(stream, DataFormats.Rtf);

                stream.Position = 0;

                // Load document from RTF stream and prepend or append clipboard content to it.
                var document = DocumentModel.Load(stream, LoadOptions.RtfDefault);
                var position = (string)e.Parameter == "prepend" ? document.Content.Start : document.Content.End;
                position.LoadFromClipboard();

                stream.Position = 0;

                // Save document to RTF stream.
                document.Save(stream, SaveOptions.RtfDefault);

                stream.Position = 0;

                // Load RTF stream into RichTextBox.
                textRange.Load(stream, DataFormats.Rtf);
            }
        }

        private void CanSave(object sender, CanExecuteRoutedEventArgs e)
        {
            if (this.richTextBox != null)
            {
                var document = this.richTextBox.Document;
                var startPosition = document.ContentStart.GetNextInsertionPosition(LogicalDirection.Forward);
                var endPosition = document.ContentEnd.GetNextInsertionPosition(LogicalDirection.Backward);
                e.CanExecute = startPosition != null && endPosition != null && startPosition.CompareTo(endPosition) < 0;
            }
            else
                e.CanExecute = false;
        }

        private void CanCut(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = this.richTextBox != null && !this.richTextBox.Selection.IsEmpty;
        }

        private void CanCopy(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = this.richTextBox != null && !this.richTextBox.Selection.IsEmpty;
        }

        private void CanPaste(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = this.richTextBox != null && this.richTextBox.IsKeyboardFocused;
        }

      
    }
}
