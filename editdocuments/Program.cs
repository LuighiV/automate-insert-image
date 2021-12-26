using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using CommandLine;
using System.IO;

namespace editdocuments
{
    internal class Program
    {
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
            public int LeftOffset { get; set; }

            [Option(
              Default = 0,
              HelpText = "Bottom offset for picture.")]
            public int BottomOffset { get; set; }

        }

        static void Main(string[] args)
        {


            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed(RunOptions)
                .WithNotParsed(HandleParseError);

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
            if (opts.Verbose)
            {
                Console.WriteLine("Start program");
            }

            var WordApp = new Word.Application();

            // If InputFolder defined use this value, otherwise use InputFiles
            IEnumerable<string> InputFiles;
            if (opts.InputFolder != null)
            {
                InputFiles = GetWordFiles(opts.InputFolder);
            }
            else
            {
                InputFiles = opts.InputFiles;
            }
            
            foreach (var FilePath in InputFiles)
            {
                // If SubFolderSave defined use this value, otherwise use FolderSave
                string FolderSave;
                if (opts.SubFolderSave != null)
                {
                    FolderSave = Path.Combine(Path.GetDirectoryName(FilePath), opts.SubFolderSave);
                }
                else
                {
                    FolderSave = opts.FolderSave;
                }

                if (!Directory.Exists(FolderSave))
                {
                    Directory.CreateDirectory(FolderSave);
                    Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(FolderSave));
                }

                RunOnFile(FilePath, 
                          opts.PicturePath, 
                          opts.PlaceHolder, 
                          opts.LeftOffset, 
                          opts.BottomOffset, 
                          WordApp, 
                          opts.WordAppVisible, 
                          FolderSave, 
                          opts.SaveFile);
            }
            

            WordApp.Quit();

            if (opts.Verbose)
            {
                Console.WriteLine("Finish program");
            }
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
            Console.WriteLine("Process file {0}", FilePath);
            var WordDocument = new WordDocument(FilePath, WordAppVisible, WordApp);

            WordDocument.GetRange(TextPlaceHolder);
            WordDocument.addPicture(PicturePath, LeftOffset, BottomOffset);
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
