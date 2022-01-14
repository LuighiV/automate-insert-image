using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace editdocuments
{
    public class PDFProcessor: Processor
    {

        public bool RunOnFile(string InputPath,
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset,
            double BottomOffset,
            double Width,
            double Height,
            string OutputPath,
            bool IsAbsolute = false,
            PageReference Reference = PageReference.bottom_left,
            int PageNumber = 1,
            bool Verbose = false)
        {

            bool success = true;
            StartProcessingFile?.Invoke(this, new TextArg(InputPath));
            if (Verbose)
                Console.WriteLine(Strings.InfoStartFile, InputPath);

            var Document = new PDFDocument(InputPath, OutputPath, Verbose);
            try
            {
                if (!IsAbsolute)
                {
                    if(Document.GetReference(TextPlaceHolder, Verbose)){
                        GotPlaceHolderPosition?.Invoke(this, new TextArg(TextPlaceHolder));
                        Document.AddPicture(PicturePath, LeftOffset, BottomOffset, Width, Height);
                    }
                    else
                    {
                        TextNotFoundInDocument?.Invoke(this, new TextArg(TextPlaceHolder));
                        success = false;
                    }
                }
                else
                {
                    if (Document.HasPage(PageNumber))
                    {
                        Document.AddPictureInPageReference(PicturePath, Reference, LeftOffset, BottomOffset, Width, Height, PageNumber);
                    }
                    else
                    {
                        PageOutOfBounds?.Invoke(this, new IntArg(PageNumber));
                        success = false;
                    }
                }

                //Console.ReadKey();
                PDFSaved?.Invoke(this, new TextArg(OutputPath));
                Document.Close();

                if (Verbose)
                    Console.WriteLine(Strings.InfoFinishFile, InputPath);

                FinishProcessingFile?.Invoke(this, new TextArg(InputPath));
            }
            catch (Exception e)
            {
                Document.Close();
                throw e;
            }
            return success;
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
            Data.FolderSave,
            Data.SubFolderSave,
            Data.IsAbsolute,
            Data.Reference,
            Data.PageNumber,
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
            string FolderSave = null,
            string SubFolderSave = null,
            bool IsAbsolute = false,
            PageReference Reference = PageReference.bottom_left,
            int PageNumber = 1,
            bool Verbose = true)
        {

            StartProcessing?.Invoke(this, EventArgs.Empty);
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoStartProgram);
                Console.WriteLine();
            }

            // If InputFolder defined use this value, otherwise use InputFiles
            if (InputFolder != null)
            {
                InputFiles = GetPDFFiles(InputFolder);
            }

            try
            {
                int index = 0;
                int count_success = 0;
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

                    string OutputPath = Path.Combine(FolderSave, Path.GetFileName(FilePath));
                    if(RunOnFile(FilePath,
                              PicturePath,
                              TextPlaceHolder,
                              LeftOffset,
                              BottomOffset,
                              Width,
                              Height,
                              OutputPath,
                              IsAbsolute,
                              Reference,
                              PageNumber,
                              Verbose))
                    {
                        count_success++;
                    }
                        
                    if (Verbose)
                    {
                        Console.WriteLine("\n\n");
                    }
                    index++;
                }

                UpdateCounter?.Invoke(this, new CounterArgs(index, totalElements));

                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoFinishProgram);
                }
                FinishProcessing?.Invoke(this, EventArgs.Empty);

                SummarySuccess?.Invoke(this, new CounterArgs(count_success, totalElements));
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        public static IEnumerable<string> GetPDFFiles(string FolderName)
        {
            return from f in Directory.EnumerateFiles(FolderName)
                   where f.EndsWith(".pdf")
                   select f;
        }

        public event EventHandler StartProcessing;
        public event EventHandler<CounterArgs> UpdateCounter;
        public event EventHandler<TextArg> StartProcessingFile;
        public event EventHandler<TextArg> GotPlaceHolderPosition;
        public event EventHandler<TextArg> TextNotFoundInDocument;
        public event EventHandler<IntArg> PageOutOfBounds;
        public event EventHandler<TextArg> PDFSaved;
        public event EventHandler<TextArg> FinishProcessingFile;
        public event EventHandler FinishProcessing;
        public event EventHandler<CounterArgs> SummarySuccess;
    }
}
