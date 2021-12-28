using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using CommandLine;
using System.IO;

namespace editdocuments
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        const int SW_SHOWMINIMIZED = 2;

        class Options
        {
            [Option('f', "filename", Default = null, HelpText = "Input files to be processed.")]
            public IEnumerable<string> InputFiles { get; set; }

            [Option('d', "inputfolder", Default = null, HelpText = "Input folder to be processed.")]
            public string InputFolder { get; set; }

            [Option('p', "picturepath", Required = true, HelpText = "Picture file to be inserted.")]
            public string PicturePath { get; set; }

            [Option('H', "placeholder", Required = true, HelpText = "Placeholder where is inserted.")]
            public string PlaceHolder { get; set; }

            
            [Option(
              Default = false,
              HelpText = "Prints all messages to standard output.")]
            public bool Verbose { get; set; }

            [Option('V', "visible",
              Default = false,
              HelpText = "Set MSWord Application visible.")]
            public bool WordAppVisible { get; set; }

            [Option('S', "foldersave",
              Default = null,
              HelpText = "Set folder to save PDF.")]
            public string FolderSave { get; set; }

            [Option('u', "subfoldersave",
              Default = null,
              HelpText = "Save in subfolder relative to document path")]
            public string SubFolderSave { get; set; }

            [Option('s', "save",
              Default = false,
              HelpText = "Save in file.")]
            public bool SaveFile { get; set; }

            [Option(
              Default = 0,
              HelpText = "Left offset for picture.")]
            public double LeftOffset { get; set; }

            [Option(
              Default = 0,
              HelpText = "Bottom offset for picture.")]
            public double BottomOffset { get; set; }

            [Option(
             Default = 200,
             HelpText = "Width for picture.")]
            public double Width { get; set; }

            [Option(
              Default = 150,
              HelpText = "Height for picture.")]
            public double Height { get; set; }

        }

        [STAThread]
        static void Main(string[] args)
        {

            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_SHOWMINIMIZED);
            //ShowWindow(handle, SW_HIDE);

            _ = Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);

            var guiForm = new GUI();
            guiForm.ShowDialog();
        }

        static void HandleParseError(IEnumerable<Error> obj)
        {
            foreach (Error error in obj)
            {
                if (error.GetType().IsAssignableFrom(typeof(MissingRequiredOptionError)))
                {
                    Console.WriteLine("Please, provide the required argument");
                    continue;
                }
                throw new Exception(error.ToString());
            }
            
        }

        static void RunOptions(Options opts)
        {
            RunProcess(opts.PicturePath,
                opts.PlaceHolder,
                opts.LeftOffset,
                opts.BottomOffset,
                opts.Width,
                opts.Height,
                opts.InputFolder,
                opts.InputFiles,
                opts.WordAppVisible,
                opts.FolderSave,
                opts.SubFolderSave,
                opts.SaveFile,
                opts.Verbose);
        }

        public static void RunProcess(
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset=0,
            double BottomOffset =0,
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
            if (Verbose)
            {
                Console.WriteLine("Start program");
            }

            var WordApp = new Word.Application();

            // If InputFolder defined use this value, otherwise use InputFiles
            if (InputFolder != null)
            {
                InputFiles = GetWordFiles(InputFolder);
            }

            foreach (var FilePath in InputFiles)
            {
                // If SubFolderSave defined use this value, otherwise use FolderSave
                if (SubFolderSave != null)
                {
                    FolderSave = Path.Combine(Path.GetDirectoryName(FilePath), SubFolderSave);
                }

                if (FolderSave!=null && !Directory.Exists(FolderSave))
                {
                    Directory.CreateDirectory(FolderSave);
                    Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(FolderSave));
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
            }


            WordApp.Quit();

            if (Verbose)
            {
                Console.WriteLine("Finish program");
            }
        }

        static void RunOnFile(string FilePath, 
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset,
            double BottomOffset,
            double Width,
            double Height,
            Word.Application WordApp=null,
            bool WordAppVisible=true,
            string FolderSave=null,
            bool SaveFile=false)
        {
            Console.WriteLine("Process file {0}", FilePath);
            var WordDocument = new WordDocument(FilePath, WordAppVisible, WordApp);

            WordDocument.GetRange(TextPlaceHolder);
            WordDocument.addPicture(PicturePath, LeftOffset, BottomOffset, Width, Height);
            WordDocument.SaveAsPDF(FolderSave);
            //Console.ReadKey();
            WordDocument.Close(SaveFile);

            Console.WriteLine("Finish processing file {0}", FilePath);
        }

        public static IEnumerable<string> GetWordFiles(string FolderName)
        {
            return from f in Directory.EnumerateFiles(FolderName)
                   where f.EndsWith(".doc") || f.EndsWith(".docx")
                   select f;
        }
    }
}
