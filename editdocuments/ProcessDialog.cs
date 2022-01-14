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
        public WordProcessor WordProcess;
        public PDFProcessor PDFProcess;

        public ProcessDialog()
        {
            InitializeComponent();

            this.IsProcessing = false;
        }

        public void  setProcessor(DocumentType type = DocumentType.Word)
        {

            if(type == DocumentType.Word)
            {
                this.WordProcess = new WordProcessor();

                this.WordProcess.StartProcessing += startProcessing;
                this.WordProcess.OpenWordProgram += openWordProgram;
                this.WordProcess.StartProcessingFile += startProcessingFile;
                this.WordProcess.GotPlaceHolderPosition += gotPlaceHolderPosition;
                this.WordProcess.UpdateCounter += updateCounter;
                this.WordProcess.PDFSaved += PDFSaved;
                this.WordProcess.FinishProcessingFile += finishProcessingFile;

                this.WordProcess.CloseWordProgram += closeWordProgram;
                this.WordProcess.FinishProcessing += finishProcessing;

                this.WordProcess.TextNotFoundInDocument += textNotFoundInDocument;
                this.WordProcess.SummarySuccess += printSummarySuccess;


                this.BackgroundProcess = this.WordProcess;
            }
            else
            {
                this.PDFProcess = new PDFProcessor();

                this.PDFProcess.StartProcessing += startProcessing;
                this.PDFProcess.StartProcessingFile += startProcessingFile;
                this.PDFProcess.GotPlaceHolderPosition += gotPlaceHolderPosition;
                this.PDFProcess.UpdateCounter += updateCounter;
                this.PDFProcess.PDFSaved += PDFSaved;
                this.PDFProcess.FinishProcessingFile += finishProcessingFile;
                this.PDFProcess.FinishProcessing += finishProcessing;
                this.PDFProcess.TextNotFoundInDocument += textNotFoundInDocument;
                this.PDFProcess.PageOutOfBounds += pageOutOfBounds;
                this.PDFProcess.SummarySuccess += printSummarySuccess;


                this.BackgroundProcess = this.PDFProcess;
            }
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

        public void textNotFoundInDocument(object sender, TextArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoTextNotFound + Environment.NewLine, e.Text));
        }

        public void pageOutOfBounds(object sender, IntArg e)
        {
            WriteTextSafe(string.Format(Strings.InfoPageOutOfBounds + Environment.NewLine, e.Value));
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

        public void printSummarySuccess(object sender, CounterArgs e)
        {
            WriteTextSafe(string.Format(Strings.InfoSummarySuccess + Environment.NewLine, e.Value, e.Total));
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

        public void RunThreadProcess(DataInfo Data)
        {
            setProcessor(Data.Type);
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
