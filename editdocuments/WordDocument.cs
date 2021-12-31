using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace editdocuments
{
    internal class WordDocument
    {
        public string DocumentPath { get; set; }

        public Word.Application WordApp;
        public Word.Document Document = new Word.Document() { };
        
        public Word.Range Range { get; set; }

        public WordDocument(string Path, bool visible = false, Word.Application WordApp = null, 
            bool Verbose = false)
        {
            this.DocumentPath = Path;
            if (WordApp == null)
            {
                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoOpenWord);
                }
                WordApp = new Word.Application() { };
            }
            this.WordApp = WordApp;
            this.WordApp.Visible = visible;
            this.Document = this.WordApp.Documents.Open(Path);

        }

        internal string SaveAsPDF(string FolderName=null, bool Verbose = false)
        {
            string Name = this.Document.Name;
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoFileName, Name);
            }

            string NewName = Path.ChangeExtension(Name, "pdf");

            if (FolderName != null)
                NewName = Path.Combine(FolderName, NewName);


            
            this.Document.SaveAs2(FileName: NewName, FileFormat: Word.WdSaveFormat.wdFormatPDF);

            if (Verbose)
            {
                Console.WriteLine(Strings.InfoSavedPDFFile, NewName);
            }

            return NewName;
        }

        internal void Close(bool save=true, bool Verbose=false)
        {
            if (save)
                this.Document.Close(Word.WdSaveOptions.wdSaveChanges);
            else
                this.Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

            if (Verbose)
            {
                Console.WriteLine(Strings.InfoSaveWordDocument, save ? Strings.InfoYes : Strings.InfoNo);
            }
            
        }

        internal void addPicture(string picture, double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
             double Height =150)
        {

            var LeftDistance = this.Range.Information[Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary];
            var shape = this.Document.Shapes.AddPicture(picture,
                Left: LeftDistance + LeftOffset,
                Top: -BottomOffset - Height,
                Width: Width,
                Height: Height,
                Anchor: this.Range);
            
        }

        public string GetRange(string text, bool Verbose=false)
        {
            var range = this.Document.Content;
            range.Find.Execute(text);
            this.Range = range;
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoPlaceholder, this.Range.Text);
            }
            return this.Range.Text;
        }
    }
}
