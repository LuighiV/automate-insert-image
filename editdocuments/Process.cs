using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace editdocuments
{
    public partial class Process : Form
    {
        public Process()
        {
            InitializeComponent();
            Console.SetOut(new ControlWriter(this.textBox1));

            Console.WriteLine("Executing in New Form");

        }

        public void setFilename(string filename)
        {
            this.label5.Text = filename;
        }


        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public  void RunGUIProcess(
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
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoStartProgram);
            }

            Console.WriteLine();
            Console.WriteLine(Strings.InfoOpenWord);
            
            var WordApp = new Word.Application();

            // If InputFolder defined use this value, otherwise use InputFiles
            if (InputFolder != null)
            {
                InputFiles = Program.GetWordFiles(InputFolder);
            }

            try
            {
                int totalElements = InputFiles.Count();
                this.label3.Text = totalElements.ToString();
                this.progressBar1.Maximum = totalElements;
                this.progressBar1.Step = 1;

                int index = 0;
                this.progressBar1.Value = index;
                this.label1.Text = index.ToString();
                foreach (var FilePath in InputFiles)
                {
                    Console.WriteLine(Strings.InfoCounterFile, index+1,totalElements);
                    Console.WriteLine();
                    setFilename(FilePath);

                    // If SubFolderSave defined use this value, otherwise use FolderSave
                    if (SubFolderSave != null)
                    {
                        FolderSave = Path.Combine(Path.GetDirectoryName(FilePath), SubFolderSave);
                    }

                    if (FolderSave != null && !Directory.Exists(FolderSave))
                    {
                        DirectoryInfo dir= Directory.CreateDirectory(FolderSave);
                        Console.WriteLine(Strings.InfoCreateDirectorySuccess, dir.FullName, dir.CreationTime);
                    }

                    Program.RunOnFile(FilePath,
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

                    Console.WriteLine("\n\n");
                    index++;
                    this.progressBar1.Value = index;
                    this.label1.Text = index.ToString();
                }


                Console.WriteLine(Strings.InfoCloseWord);
                Console.WriteLine();
                WordApp.Quit();

                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoFinishProgram);
                }
            }
            catch (Exception e)
            {
                WordApp.Quit();
                throw e;
            }

        }
    }
}
