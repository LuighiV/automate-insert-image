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

        /// <summary>
        /// Class with the options to pass to the command line editordocuments
        /// </summary>
        class Options
        {
            [Option('t', "type", Default = DocumentType.Word, HelpText = "Document type.")]
            public DocumentType Type { get; set; }

            [Option('A', "is-absolute", Default = false, HelpText = "The reference is absolute to the page.")]
            public bool IsAbsolute { get; set; }

            [Option('N', "page", Default = 1, HelpText = "Page number (valid for page reference type).")]
            public int PageNumber { get; set; }

            [Option('r', "page-reference", Default = PageReference.bottom_left, HelpText = "Page reference.")]
            public PageReference Reference { get; set; }

            [Option('f', "filename", Default = null, HelpText = "Input files to be processed.")]
            public IEnumerable<string> InputFiles { get; set; }

            [Option('d', "inputfolder", Default = null, HelpText = "Input folder to be processed.")]
            public string InputFolder { get; set; }

            [Option('p', "picturepath", Default = null, HelpText = "Picture file to be inserted.")]
            public string PicturePath { get; set; }

            [Option('H', "placeholder", Default = null, HelpText = "Placeholder where is inserted.")]
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
            //ShowWindow(handle, SW_SHOWMINIMIZED);
            ShowWindow(handle, SW_HIDE);

            if (args.Length == 0)
            {
                var guiForm = new GUI();
                guiForm.ShowDialog();
            }
            else
            {
                var parser = new Parser(with => with.HelpWriter = Console.Error);
                _ = parser.ParseArguments<Options>(args)
                    .WithParsed(RunOptions)
                    .WithNotParsed(HandleParseError);
            }

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
                // When trigger --help or --version, it returns errors
                // Console.WriteLine(error.ToString());
            }
            
        }
        /// <summary>
        /// Run the program with the options obtained from the command line
        /// </summary>
        /// <param name="opts"> Command line options to pass to the software</param>
        static void RunOptions(Options opts)
        {
            if(opts.Type == DocumentType.Word)
            {
                if (opts.Verbose)
                {
                    Console.WriteLine("Document type is set to Word");
                }
                var processor = new WordProcessor();
                processor.RunProcess(opts.PicturePath,
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

            if (opts.Type == DocumentType.PDF)
            {
                if (opts.Verbose)
                {
                    Console.WriteLine("Document type is set to PDF");
                }
                var processor = new PDFProcessor();
                processor.RunProcess(opts.PicturePath,
                    opts.PlaceHolder,
                    opts.LeftOffset,
                    opts.BottomOffset,
                    opts.Width,
                    opts.Height,
                    opts.InputFolder,
                    opts.InputFiles,
                    opts.FolderSave,
                    opts.SubFolderSave,
                    opts.IsAbsolute,
                    opts.Reference,
                    opts.PageNumber,
                    opts.Verbose);
            }

        }
    }
}
