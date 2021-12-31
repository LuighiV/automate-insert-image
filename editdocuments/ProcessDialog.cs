using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace editdocuments
{
    public partial class ProcessDialog : Form
    {

        public bool IsProcessing;
        public Processor BackgroundProcess;
        public ProcessDialog()
        {
            InitializeComponent();
            //Console.SetOut(new ControlWriter(this.textBox1));

            //Console.WriteLine("Executing in New Form");
            this.BackgroundProcess = new Processor();

            this.BackgroundProcess.StartProcessing += startProcessing;
            this.BackgroundProcess.OpenWordProgram += openWordProgram;
            this.BackgroundProcess.StartProcessingFile += startProcessingFile;
            this.BackgroundProcess.GotPlaceHolderPosition += gotPlaceHolderPosition;
            this.BackgroundProcess.UpdateCounter += updateCounter;
            this.BackgroundProcess.PDFSaved += PDFSaved;
            this.BackgroundProcess.FinishProcessingFile += finishProcessingFile;

            this.BackgroundProcess.CloseWordProgram += closeWordProgram;
            this.BackgroundProcess.FinishProcessing += finishProcessing;
            //this.ControlBox = false;  // Hides controls of window

            this.IsProcessing = false;
        }

        public void successFullFinished()
        {
            MessageBox.Show(Strings.ExecutionCompletedDefaultMessage,
                    Strings.ExecutionCompletedTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            buttonSafeEnable(true);
            this.IsProcessing = false;
        }

        public void buttonSafeEnable(bool enable=true)
        {
            if (this.button1.InvokeRequired)
            {
                Action safeWrite = delegate { buttonSafeEnable(enable); };
                this.button1.Invoke(safeWrite);
            }
            else
            {
                this.button1.Enabled= enable;
            }
        }


        public void startProcessing(object sender, EventArgs e)
        {
            WriteTextSafe(Strings.InfoStartProgram + Environment.NewLine);
            WriteTextSafe(Environment.NewLine);
        }

        public void openWordProgram(object sender, EventArgs e)
        {
            WriteTextSafe(Strings.InfoOpenWord + Environment.NewLine);
            WriteTextSafe(Environment.NewLine);
        }

        public void updateCounter(object sender, CounterArgs e)
        {
            if(e.Value ==0)
            {
                WriteTextSafe(string.Format(Strings.InfoStartProcessingFiles + Environment.NewLine, 
                    e.Total));
                WriteTextSafe(Environment.NewLine);
                SetupProgressBarSafe(e.Total, 1, this.progressBar1);
            }

            WriteLabelSafe(e.Value.ToString(), this.label1);
            WriteLabelSafe(e.Total.ToString(), this.label3);
            UpdateProgressBarSafe(e.Value, this.progressBar1);

            if(e.Value < e.Total)
            {
                int currentFileNumber = e.Value + 1;
                WriteTextSafe(string.Format(Strings.InfoCounterFile + Environment.NewLine,
                    currentFileNumber.ToString(),
                    e.Total.ToString()));
                WriteTextSafe(Environment.NewLine);
            }
            else
            {
                WriteTextSafe(Strings.InfoFinishProcessingFiles + Environment.NewLine);
                WriteTextSafe(Environment.NewLine);
            }

        }

        public void startProcessingFile(object sender, TextArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoStartFile + Environment.NewLine, e.Text));
            WriteLabelSafe( e.Text,this.label5);
        }

        public void gotPlaceHolderPosition(object sender, TextArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoPlaceholder + Environment.NewLine, e.Text));
        }

        public void PDFSaved(object sender, TextArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoSavedPDFFile + Environment.NewLine, e.Text));
        }


        public void finishProcessingFile(object sender, TextArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoFinishFile + Environment.NewLine, e.Text));
            WriteTextSafe(Environment.NewLine);
        }

        public void closeWordProgram(object sender, EventArgs e)
        {

            WriteTextSafe(Strings.InfoCloseWord + Environment.NewLine);
            WriteTextSafe(Environment.NewLine);
        }

        public void finishProcessing(object sender, EventArgs e)
        {
            WriteTextSafe(Strings.InfoFinishProgram + Environment.NewLine);
            WriteTextSafe(Environment.NewLine);
        }

        public void WriteTextSafe(string text)
        {
            if (this.textBox1.InvokeRequired)
            {
                Action safeWrite = delegate { WriteTextSafe(text); };
                this.textBox1.Invoke(safeWrite);
            }
            else{
                this.textBox1.AppendText(text);
            }
        }

        public void WriteLabelSafe(string text, Label label)
        {
            if (label.InvokeRequired)
            {
                Action safeWrite = delegate { WriteLabelSafe(text, label); };
                label.Invoke(safeWrite);
            }
            else
            {
                label.Text = text;
            }
        }

        public void SetupProgressBarSafe(int maximum, int step, ProgressBar progressbar)
        {
            if (progressbar.InvokeRequired)
            {
                Action safeWrite = delegate { SetupProgressBarSafe(maximum, step, progressbar); };
                progressbar.Invoke(safeWrite);
            }
            else
            {
                progressbar.Maximum = maximum;
                progressbar.Step = step;
            }
        }

        public void UpdateProgressBarSafe(int value, ProgressBar progressbar)
        {
            if (progressbar.InvokeRequired)
            {
                Action safeWrite = delegate { UpdateProgressBarSafe(value, progressbar); };
                progressbar.Invoke(safeWrite);
            }
            else
            {
                progressbar.Value = value;
            }
        }

        public void HandleError(Exception e)
        {
            Console.WriteLine(e.ToString());
            MessageBox.Show(e.Message,
                Strings.ExecutionErrorTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        public void setFilename(string filename)
        {
            this.label5.Text = filename;
        }


        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void RunThreadProcess(DataInfo Data)
        {
            var threadParameters = new ThreadStart(delegate
            {
                try
                {
                    this.IsProcessing = true;
                    this.BackgroundProcess.RunProcess(Data);
                }
                catch(Exception e)
                {
                    HandleError(e);
                }
            
            });
            threadParameters += () => { successFullFinished(); };
            
            var thread = new Thread(threadParameters);
            this.button1.Enabled = false;
            thread.Start();
            
        }

        public void RunGUIProcess(DataInfo Data, bool Verbose=false)
        {
            GUnits tmpUnit = Data.Unit;
            Data.Unit = GUnits.point;
            RunGUIProcess(
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

        public  void RunGUIProcess(
            string PicturePath,
            string TextPlaceHolder,
            double LeftOffset = 0,
            double BottomOffset = 0,
            double Width = 200,
            double Height = 150,
            string InputFolder = null,
            IEnumerable<string> InputFiles = null,
            bool WordAppVisible = false,
            string FolderSave = null,
            string SubFolderSave = null,
            bool SaveFile = false,
            bool Verbose = true)
        {
            if (Verbose)
            {
                Console.WriteLine(Strings.InfoStartProgram);
                Console.WriteLine();
                Console.WriteLine(Strings.InfoOpenWord);
            }
            
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
                    if (Verbose)
                    {
                        Console.WriteLine(Strings.InfoCounterFile, index + 1, totalElements);
                        Console.WriteLine();
                    }

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
                    
                    if (Verbose)
                    {
                        Console.WriteLine(Environment.NewLine + Environment.NewLine);
                    }
                    
                    index++;
                    this.progressBar1.Value = index;
                    this.label1.Text = index.ToString();
                }

                if (Verbose)
                {
                    Console.WriteLine(Strings.InfoCloseWord);
                    Console.WriteLine();
                }

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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ProcessDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.IsProcessing)
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }
    }
}
