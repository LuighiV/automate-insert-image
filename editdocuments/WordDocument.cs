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

        public WordDocument(string Path, bool visible = false, Word.Application WordApp = null)
        {
            this.DocumentPath = Path;
            if (WordApp == null)
            {
                WordApp = new Word.Application() { };
            }
            this.WordApp = WordApp;
            this.WordApp.Visible = visible;
            this.Document = this.WordApp.Documents.Open(Path);

        }

        internal void SaveAsPDF(string FolderName=null)
        {
            string Name = this.Document.Name;
            string NewName = Path.ChangeExtension(Name, "pdf");

            if (FolderName != null)
                NewName = Path.Combine(FolderName, NewName);

            Console.WriteLine("Current filename: {0}", Name);
            this.Document.SaveAs2(FileName: NewName, FileFormat: Word.WdSaveFormat.wdFormatPDF);
        }

        internal void Close(bool save=true)
        {
            if (save)
                this.Document.Close(Word.WdSaveOptions.wdSaveChanges);
            else
                this.Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
        }

        internal void addPicture(string picture, int LeftOffset = 0, int BottomOffset = 0, int Width = 200,
             int Height=150)
        {
            //var borders = this.Range.Borders;
            //Console.WriteLine(borders.DistanceFromLeft);
            var shape = this.Document.Shapes.AddPicture(picture,
                Left: LeftOffset,
                Top: BottomOffset - Height,
                Width: Width,
                Height: Height,
                Anchor: this.Range);
            
        }

        public Word.Range GetRange(string text)
        {
            var range = this.Document.Content;
            range.Find.Execute(text);
            this.Range = range;
            Console.WriteLine(this.Range.Text);
            return range;
        }
    }
}
