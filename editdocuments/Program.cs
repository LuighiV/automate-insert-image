using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using CommandLine;

namespace editdocuments
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start program");
            string FilePath = @"D:\Luighi\Documents\SITUACIÓN FINAL 3 años INFORME DEL PROGRESO DEL-LA ESTUDIANTE 2021.doc";
            string PicturePath = @"D:\Luighi\Documents\firmaLuighiViton.png";
            var WordApp = new Word.Application();

            RunOnFile(FilePath, PicturePath, "DIRECTORA (e)", 280, 10, WordApp, false, null, false);

            WordApp.Quit();

            Console.WriteLine("Finish program");
        }

        static void RunOnFile(string FilePath, 
            string PicturePath,
            string TextPlaceHolder,
            int LeftOffset,
            int BottomOffset,
            Word.Application WordApp=null,
            bool WordAppVisible=true,
            string FolderSave=null,
            bool SaveFile=false)
        {
            var WordDocument = new WordDocument(FilePath, WordAppVisible, WordApp);

            WordDocument.GetRange(TextPlaceHolder);
            WordDocument.addPicture(PicturePath, LeftOffset, BottomOffset);
            WordDocument.SaveAsPDF(FolderSave);
            Console.ReadKey();
            WordDocument.Close(SaveFile);
        }
    }
}
