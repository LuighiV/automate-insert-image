﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace editdocuments
{
    public class Processor
    {

        public Processor()
        {

        }

        public void RunOnFile(string FilePath,
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset,
            double BottomOffset,
            double Width,
            double Height,
            Word.Application WordApp = null,
            bool WordAppVisible = true,
            string FolderSave = null,
            bool SaveFile = false,
            bool Verbose = false)
        {

            StartProcessingFile?.Invoke(this, new TextArg(FilePath));
            if(Verbose)
                Console.WriteLine(Strings.InfoStartFile, FilePath);

            var WordDocument = new WordDocument(FilePath, WordAppVisible, WordApp);
            try
            {
                string textRange = WordDocument.GetRange(TextPlaceHolder);
                GotPlaceHolderPosition?.Invoke(this, new TextArg(textRange));

                WordDocument.addPicture(PicturePath, LeftOffset, BottomOffset, Width, Height);
                string pdffilename = WordDocument.SaveAsPDF(FolderSave);
                PDFSaved?.Invoke(this, new TextArg(pdffilename));
                //Console.ReadKey();
                WordDocument.Close(SaveFile);

                if(Verbose)
                    Console.WriteLine(Strings.InfoFinishFile, FilePath);

                FinishProcessingFile?.Invoke(this, new TextArg(FilePath));
            }
            catch (Exception e)
            {
                WordDocument.Close(false);
                throw e;
            }

        }

        public void RunProcess(DataInfo Data, bool Verbose = false)
        {
            GUnits tmpUnit = Data.Unit;
            Data.Unit = GUnits.point;
            RunProcess(
            Data.PicturePath,
            Data.TextPlaceHolder,
            Data.LeftOffset,
            Data.BottomOffset,
            Data.Width,
            Data.Height,
            Data.FolderPath,
            Data.InputFilePaths,
            Data.WordAppVisible,
            Data.FolderSave,
            Data.SubFolderSave,
            Data.SaveFile,
            Verbose);

            Data.Unit = tmpUnit;
        }

        public void RunProcess(
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset = 0,
            double BottomOffset = 0,
            double Width = 200,
            double Height = 150,
            string InputFolder = null,
            IEnumerable<string> InputFiles = null,
            bool WordAppVisible = true,
            string FolderSave = null,
            string SubFolderSave = null,
            bool SaveFile = false,
            bool Verbose = true)
        {

            StartProcessing?.Invoke(this, EventArgs.Empty);
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoStartProgram);
                Console.WriteLine();
                Console.WriteLine(Strings.InfoOpenWord);
            }

            OpenWordProgram?.Invoke(this, EventArgs.Empty);

            var WordApp = new Word.Application();

            // If InputFolder defined use this value, otherwise use InputFiles
            if (InputFolder != null)
            {
                InputFiles = GetWordFiles(InputFolder);
            }

            try
            {
                int index = 0;
                int totalElements = InputFiles.Count();
                foreach (var FilePath in InputFiles)
                {
                    UpdateCounter?.Invoke(this, new CounterArgs(index, totalElements));
                    if (Verbose)
                    {
                        Console.WriteLine(Strings.InfoCounterFile, index + 1, totalElements);
                        Console.WriteLine();
                    }
                    // If SubFolderSave defined use this value, otherwise use FolderSave
                    if (SubFolderSave != null)
                    {
                        FolderSave = Path.Combine(Path.GetDirectoryName(FilePath), SubFolderSave);
                    }

                    if (FolderSave != null)
                    {
                        if (!Directory.Exists(FolderSave))
                        {
                            DirectoryInfo dir = Directory.CreateDirectory(FolderSave);
                            Console.WriteLine(Strings.InfoCreateDirectorySuccess, dir.FullName, dir.CreationTime);
                        }
                    }
                    else
                    {
                        FolderSave = Path.GetDirectoryName(FilePath);
                    }

                    RunOnFile(FilePath,
                              PicturePath,
                              TextPlaceHolder,
                              LeftOffset,
                              BottomOffset,
                              Width,
                              Height,
                              WordApp,
                              WordAppVisible,
                              FolderSave,
                              SaveFile);
                    if (Verbose)
                    {
                        Console.WriteLine("\n\n");
                    }
                    index++;
                }

                UpdateCounter?.Invoke(this, new CounterArgs(index, totalElements));

                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoCloseWord);
                    Console.WriteLine();
                }

                CloseWordProgram?.Invoke(this, EventArgs.Empty);
                WordApp.Quit();

                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoFinishProgram);
                }
                FinishProcessing?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception e)
            {
                WordApp.Quit();
                throw e;
            }

        }
        public static IEnumerable<string> GetWordFiles(string FolderName)
        {
            return from f in Directory.EnumerateFiles(FolderName)
                   where f.EndsWith(".doc") || f.EndsWith(".docx")
                   select f;
        }

        public event EventHandler StartProcessing;
        public event EventHandler OpenWordProgram;
        public event EventHandler<CounterArgs> UpdateCounter;
        public event EventHandler<TextArg> StartProcessingFile;
        public event EventHandler<TextArg> GotPlaceHolderPosition;
        public event EventHandler<TextArg> PDFSaved;
        public event EventHandler<TextArg> FinishProcessingFile;
        public event EventHandler CloseWordProgram;
        public event EventHandler FinishProcessing;
    }

    public class TextArg: EventArgs
    {
        public TextArg(string text)
        {
            this.Text = text;
        }
        public string Text { get; set; }
    }

    public class CounterArgs: EventArgs
    {
        public CounterArgs(int value, int total)
        {
            this.Value = value;
            this.Total = total;
        }

        public int Total { get; set; }
        public int Value { get; set; }
    }
}
