using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;

namespace editdocuments
{
    public partial class GUI : Form
    {

        public DataInfo Data = new DataInfo(fromSettings: true);

        public bool SelectedFiles = true;
        public bool KeepAspect = true;
        public bool IgnoreTextChanged = false;
        public Image CurrentImage = null;

        public CultureInfo cultureInfo;

        public bool SettingsHasChanged = false;

        public string CurrentStatus;

        public GUI()
        {
            InitializeComponent();
            this.unitComboBox.DataSource = Globals.AvailableUnits;
            this.unitComboBox.DisplayMember = "Literal";
            this.cultureInfo = CultureInfo.CurrentCulture;

            if(this.cultureInfo.Parent.Name == "es" | this.cultureInfo.Name == "es")
            {
                this.spanishToolStripMenuItem.Checked = true;
                this.englishToolStripMenuItem.Checked = false;

            }
            else
            {
                this.spanishToolStripMenuItem.Checked = false;
                this.englishToolStripMenuItem.Checked = true;
            }

            this.relativeToPageOptionButton.Enabled = !this.wordDocumentOptionButton.Checked;
            this.sameFolderButtonOption.Enabled = this.wordDocumentOptionButton.Checked;
            enableTextReference(this.relativeToTextOptionButton.Checked);
            enablePageReference(!this.relativeToTextOptionButton.Checked);

            if (IsValidPicturePath())
            {
                this.CurrentImage = new Bitmap(this.picturePathTextBox.Text);
                fillPictureBox(this.picturePreviewBox, this.CurrentImage);
            }

            SetStatus(Strings.InfoStatusReady);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.Data.InputPath = this.inputPathTextBox.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.Data.IsFilesSelected)
            {
                SetStatus(Strings.InfoStatusFileExplorerOpen);
                if (this.Data.Type == DocumentType.Word)
                    this.openWordFileDialog.ShowDialog();
                else
                    this.openPDFFileDialog.ShowDialog();
                SetStatus(Strings.InfoStatusReady);
            }
            else
            {
                SetStatus(Strings.InfoStatusFolderExplorerOpen);
                DialogResult result =  this.folderBrowserDialog1.ShowDialog();
                SetStatus(Strings.InfoStatusReady);
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(this.folderBrowserDialog1.SelectedPath))
                {
                    this.Data.InputPath = this.folderBrowserDialog1.SelectedPath;
                    this.inputPathTextBox.Text = this.Data.InputPath;   
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.Data.IsFilesSelected= true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            this.Data.IsFilesSelected = false;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            this.inputPathTextBox.Text = String.Join(",", this.openWordFileDialog.FileNames);
            this.Data.InputPath = this.inputPathTextBox.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetStatus(Strings.InfoStatusFileExplorerOpen);
            this.openImageFileDialog.ShowDialog();
            SetStatus(Strings.InfoStatusReady);
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            this.IgnoreTextChanged = true;

            this.Data.PicturePath = this.openImageFileDialog.FileName;
            this.picturePathTextBox.Text = this.Data.PicturePath;

            this.CurrentImage = new Bitmap(this.Data.PicturePath);
            this.imageWidthNumeric.Value = this.CurrentImage.Width;
            this.imageHeightNumeric.Value = this.CurrentImage.Height;

            fillPictureBox(this.picturePreviewBox, this.CurrentImage);

            InitiateDimensions();
            ConvertUnits();
            UpdateDimensionsInUI();

            this.picturePathTextBox.Enabled = true;

            this.IgnoreTextChanged = false;
        }

        private void InitiateDimensions()
        {
            this.Data.LeftOffset = (double)this.imageLeftOffsetNumeric.Value;
            this.Data.BottomOffset = (double)this.imageBottomOffsetNumeric.Value;

            this.Data.Unit = GUnits.pixel;
            this.Data.Width = (double)this.imageWidthNumeric.Value;
            this.Data.Height = (double)this.imageHeightNumeric.Value;
            
        }

        private void ConvertUnits()
        {
            UnitStruct currentUnit = (UnitStruct) this.unitComboBox.SelectedItem;
            this.Data.Unit = currentUnit.Unit;
        }

        private void UpdateDimensionsInUI()
        {
            this.imageWidthNumeric.Value = (decimal)this.Data.Width;
            this.imageHeightNumeric.Value = (decimal)this.Data.Height;
            this.imageLeftOffsetNumeric.Value = (decimal)this.Data.LeftOffset;
            this.imageBottomOffsetNumeric.Value = (decimal)this.Data.BottomOffset;
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.IgnoreTextChanged = true;
            ConvertUnits();
            UpdateDimensionsInUI();
            this.IgnoreTextChanged = false;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.Data.PicturePath= this.picturePathTextBox.Text;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.KeepAspect = this.keepAspectCheckBox.Checked;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;

            double aspectRatio = this.Data.Width / this.Data.Height;

            this.Data.Width = (double) this.imageWidthNumeric.Value;
            if (this.KeepAspect)
            {

                this.IgnoreTextChanged = true;
                this.Data.Height = this.Data.Width / aspectRatio;
                this.imageHeightNumeric.Value = (decimal)this.Data.Height;
                this.imageWidthNumeric.Value = (decimal)this.Data.Width; //required to prevent not save data en GUI
                this.IgnoreTextChanged = false;
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;

            double aspectRatio = this.Data.Width / this.Data.Height;

            this.Data.Height = (double)this.imageHeightNumeric.Value;
            if (this.KeepAspect)
            {
                this.IgnoreTextChanged = true;
                this.Data.Width = this.Data.Height * aspectRatio;
                this.imageWidthNumeric.Value = (decimal)this.Data.Width;
                this.imageHeightNumeric.Value = (decimal)this.Data.Height; //required to prevent not save data en GUI
                this.IgnoreTextChanged = false;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            this.Data.TextPlaceHolder = this.placeHolderTextBox.Text;
        }

        private void SetStatus(string text)
        {
            this.toolStripStatus.Text = text;
            this.CurrentStatus = text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetStatus(string.Format(Strings.InfoStatusStartProcessing));

            if (!ValidateChildren(ValidationConstraints.Enabled))
            {
                SetStatus(string.Format(Strings.InfoStatusValidationError));
                return;
            }

            this.toolStripProgressBar.Value = 30;

            try
            {
                this.Enabled = false;
                Application.UseWaitCursor = true;
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();
                var processForm = new ProcessDialog();
                processForm.CloseForm += ProcessForm_CloseForm;
                processForm.Show();

                processForm.RunThreadProcess(this.Data);
                SetStatus(string.Format(Strings.InfoStatusNewDialogProcessing));
                this.toolStripProgressBar.Value = 100;

            }
            catch(Exception error)
            {
                Console.WriteLine(error.ToString());
                MessageBox.Show(error.ToString(),
                    Strings.ExecutionErrorTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.Enabled = true;
                this.Cursor = Cursors.Default;
                Application.UseWaitCursor = false;

            }

        }

        private void ProcessForm_CloseForm(object sender, EventArgs e)
        {
            SetStatus(string.Format(Strings.InfoStatusReady));
            this.toolStripProgressBar.Value = 0;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;
            this.Data.LeftOffset = (double)this.imageLeftOffsetNumeric.Value;
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;
            this.Data.BottomOffset = (double)this.imageBottomOffsetNumeric.Value;
        }

        // Funtions to verify data

        private bool IsValidSelectedPath()
        {
            return (!String.IsNullOrEmpty(inputPathTextBox.Text));
        }

        private bool IsValidPicturePath()
        {
            return (!String.IsNullOrEmpty(picturePathTextBox.Text));
        }

        private bool IsValidPlaceHolder()
        {
            return (!String.IsNullOrEmpty(placeHolderTextBox.Text));
        }

        private bool IsValidSubFolder()
        {
            return (!String.IsNullOrEmpty(subFolderTextBox.Text));
        }


        //Reference https://stackoverflow.com/a/56119473
        static public void fillPictureBox(PictureBox pbox, Image bmp)
        {
            pbox.SizeMode = PictureBoxSizeMode.Normal;
            bool source_is_wider = (float)bmp.Width / bmp.Height > (float)pbox.Width / pbox.Height;

            var resized = new Bitmap(pbox.Width, pbox.Height);
            var g = Graphics.FromImage(resized);
            var dest_rect = new Rectangle(0, 0, pbox.Width, pbox.Height);
            Rectangle src_rect;

            if (source_is_wider)
            {
                float size_ratio = (float)pbox.Width / bmp.Width;
                int sample_height = (int)(pbox.Height / size_ratio);
                src_rect = new Rectangle(0, (bmp.Height - sample_height) / 2, bmp.Width, sample_height);
            }
            else
            {
                float size_ratio = (float)pbox.Height / bmp.Height;
                int sample_width = (int)(pbox.Width / size_ratio);
                src_rect = new Rectangle((bmp.Width - sample_width) / 2, 0, sample_width, bmp.Height);
            }

            g.DrawImage(bmp, dest_rect, src_rect, GraphicsUnit.Pixel);
            g.Dispose();

            pbox.Image = resized;
        }


        private void textBox1_Validating(object sender, CancelEventArgs e)
        {

            if (IsValidSelectedPath())
            {
                inputPathErrorProvider.SetError(this.inputPathTextBox, String.Empty);
                e.Cancel = false;
            }
            else
            {
                inputPathErrorProvider.SetError(this.inputPathTextBox, Strings.TextValidationFilePath);
                e.Cancel = this.generateButton.Focused;
            }
            
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
            if (IsValidPicturePath())
            {
                picturePathErrorProvider.SetError(this.picturePathTextBox, String.Empty);
                e.Cancel = false;
            }
            else
            {
                picturePathErrorProvider.SetError(this.picturePathTextBox, Strings.TextValidationPicturePath);
                e.Cancel = this.generateButton.Focused;
            }
        }

        private void placeHolderTextBox_Validating(object sender, CancelEventArgs e)
        {

            if (IsValidPlaceHolder() | this.relativeToPageOptionButton.Checked)
            {
                placeHolderErrorProvider.SetError(this.placeHolderTextBox, String.Empty);
                e.Cancel = false;
            }
            else
            {
                placeHolderErrorProvider.SetError(this.placeHolderTextBox, Strings.TextValidationPlaceholder);
                e.Cancel = this.generateButton.Focused;
            }
        }

        private void subFolderButtonOption_CheckedChanged(object sender, EventArgs e)
        {
            this.subFolderErrorProvider.SetError(this.subFolderTextBox, String.Empty);
            this.Data.IsSubFolderSelected = true;
            enableSubFolderTextBox(this.subFolderButtonOption.Checked);
        }

        private void sameFolderButtonOption_CheckedChanged(object sender, EventArgs e)
        {
            this.Data.IsSubFolderSelected = false;
        }

        private void subFolderTextBox_TextChanged(object sender, EventArgs e)
        {
            this.Data.SubFolderSave = this.subFolderTextBox.Text;
        }

        private void englishToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeToCulture(new CultureInfo("en-US"));
            this.englishToolStripMenuItem.Checked = true;
            this.spanishToolStripMenuItem.Checked = false;
        }

        private void changeToCulture(CultureInfo cultureInfo)
        {
            Thread.CurrentThread.CurrentUICulture = cultureInfo;
            ComponentResourceManager resources = new ComponentResourceManager(typeof(GUI));
            resources.ApplyResources(this, "$this");
            applyResources(resources, this.Controls);
        }

        // reference https://stackoverflow.com/a/7558253
        private void applyResources(ComponentResourceManager resources, Control.ControlCollection ctls)
        {
            foreach (Control ctl in ctls)
            {
                resources.ApplyResources(ctl, ctl.Name);
                applyResources(resources, ctl.Controls);
            }
        }

        private void spanishToolStripMenuItem_Click(object sender, EventArgs e)
        {
            changeToCulture(new CultureInfo("es"));
            this.englishToolStripMenuItem.Checked = false;
            this.spanishToolStripMenuItem.Checked = true;
        }

        private void saveCurrentValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();

            //Console.WriteLine(Properties.Settings.Default.PicturePath);
            this.saveCurrentValuesToolStripMenuItem.Enabled = false;
            this.SettingsHasChanged = false;
        }

        private void resetToDefaultsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
            this.picturePreviewBox.Image = null; 
            //Console.WriteLine("Reset settings");
        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var infoDialog = new InfoDialog(Strings.TitleHowToUse, Files.howToUse);
            infoDialog.Show();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var infoDialog = new InfoDialog(Strings.TitleAbout, Files.about);
            infoDialog.Show();
        }

        private void GUI_Load(object sender, EventArgs e)
        {
            Properties.Settings.Default.PropertyChanged += Default_PropertyChanged;
            Properties.Settings.Default.SettingsSaving += Default_SettingsSaving;
        }

        private void Default_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            this.SettingsHasChanged = true;
            this.saveCurrentValuesToolStripMenuItem.Enabled = true;
            //Console.WriteLine($"{e.PropertyName}");
        }

        private void Default_SettingsSaving(object sender, CancelEventArgs e)
        {
            //Console.WriteLine("Saving settings");
        }

        private void GUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SettingsHasChanged)
            {
                DialogResult result = MessageBox.Show(Strings.QuerySaveCurrentValues,
                    Strings.SaveCurrentValues,
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Information);

                if(result == DialogResult.Yes)
                {
                    Properties.Settings.Default.Save();
                } 
                else if (result == DialogResult.No)
                {

                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void pdfDocumentOptionButton_Click(object sender, EventArgs e)
        {
            this.pdfDocumentOptionButton.Checked = true;
            this.subFolderButtonOption.Checked = true;
            this.Data.Type = DocumentType.PDF;
        }

        private void wordDocumentOptionButton_Click(object sender, EventArgs e)
        {
            this.wordDocumentOptionButton.Checked = true;
            this.relativeToPageOptionButton.Checked = false;
            this.relativeToTextOptionButton.Checked = true;
            this.Data.Type = DocumentType.Word;
        }

        private void topLeftOptionButton_Click(object sender, EventArgs e)
        {
            this.topLeftOptionButton.Checked = true;
            this.Data.Reference = PageReference.top_left;
        }

        private void topRightOptionButton_Click(object sender, EventArgs e)
        {
            this.topRightOptionButton.Checked = true;
            this.Data.Reference = PageReference.top_right;
        }

        private void bottomLeftOptionButton_Click(object sender, EventArgs e)
        {
            this.bottomLeftOptionButton.Checked = true;
            this.Data.Reference = PageReference.bottom_left;
        }

        private void bottomRightOptionButton_Click(object sender, EventArgs e)
        {
            this.bottomRightOptionButton.Checked = true;
            this.Data.Reference = PageReference.bottom_right;
        }

        private void relativeToTextOptionButton_Click(object sender, EventArgs e)
        {
            this.relativeToTextOptionButton.Checked = true;
            this.Data.IsAbsolute = false;
        }

        private void relativeToPageOptionButton_Click(object sender, EventArgs e)
        {
            this.relativeToPageOptionButton.Checked = true;
            this.Data.IsAbsolute = true;
        }

        private void pageNumberNumeric_ValueChanged(object sender, EventArgs e)
        {
            this.Data.PageNumber = (int)this.pageNumberNumeric.Value;
        }

        private void enableTextReference(bool enable)
        {
            this.placeHolderTextBox.Enabled = enable;
            this.placeHolderLabel.Enabled = enable;
        }

        private void enablePageReference(bool enable)
        {
            this.pageReferenceLabel.Enabled = enable;
            this.pageNumberLabel.Enabled = enable;
            this.pageNumberNumeric.Enabled = enable;
            this.relativeToGroupBox.Enabled = enable;
        }

        private void enableSubFolderTextBox(bool enable)
        {
            this.subFolderTextBox.Enabled = enable;
        }

        private void wordDocumentOptionButton_CheckedChanged(object sender, EventArgs e)
        {
            this.relativeToPageOptionButton.Enabled = !this.wordDocumentOptionButton.Checked;
        }

        private void pdfDocumentOptionButton_CheckedChanged(object sender, EventArgs e)
        {
            this.sameFolderButtonOption.Enabled = this.wordDocumentOptionButton.Checked;
        }

        private void relativeToTextOptionButton_CheckedChanged(object sender, EventArgs e)
        {
            this.placeHolderErrorProvider.SetError(this.placeHolderTextBox, String.Empty);
            enableTextReference(this.relativeToTextOptionButton.Checked);
            enablePageReference(!this.relativeToTextOptionButton.Checked);
        }

        private void openPDFFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            this.inputPathTextBox.Text = String.Join(",", this.openPDFFileDialog.FileNames);
            this.Data.InputPath = this.inputPathTextBox.Text;
        }

        private void subFolderTextBox_Validating(object sender, CancelEventArgs e)
        {
            if (IsValidSubFolder() | !this.subFolderButtonOption.Checked)
            {
                subFolderErrorProvider.SetError(this.subFolderTextBox, String.Empty);
                e.Cancel = false;
            }
            else
            {
                subFolderErrorProvider.SetError(this.subFolderTextBox, Strings.TextValidationSubFolder);
                e.Cancel = this.generateButton.Focused;
            }
        }

        private void pictureBox1_SizeChanged(object sender, EventArgs e)
        {
            if (this.CurrentImage != null & this.picturePreviewBox.Width >0 & this.picturePreviewBox.Height > 0)
            {
                fillPictureBox(this.picturePreviewBox, this.CurrentImage);
            }
        }

        private void wordDocumentOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpWordSelection);
        }

        private void wordDocumentOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void pdfDocumentOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpPDFSelection);
        }

        private void pdfDocumentOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void filesOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpInputPathFiles);
        }

        private void filesOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void folderOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpInputPathFolder);
        }

        private void folderOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void inputPathTextBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpInputPath);
        }

        private void inputPathTextBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void exploreInputPathButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpInputPathExplore);
        }

        private void exploreInputPathButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void picturePathTextBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpPicturePath);
        }

        private void picturePathTextBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void explorePicturePathButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpPicturePathExplore);
        }

        private void explorePicturePathButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void keepAspectCheckBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpKeepAspectRatio);
        }

        private void keepAspectCheckBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void unitComboBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpCurrentUnit);
        }

        private void unitComboBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void picturePreviewBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpImagePreview);
        }

        private void picturePreviewBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void relativeToTextOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpRelativeToText);
        }

        private void relativeToTextOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void relativeToPageOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpRelativeToPage);
        }

        private void relativeToPageOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void placeHolderTextBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpPlaceHolder);
        }

        private void placeHolderTextBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void topLeftOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpTopLeft);
        }

        private void topLeftOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void bottomLeftOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpBottomLeft);
        }

        private void bottomLeftOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void topRightOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpTopRight);
        }

        private void topRightOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void bottomRightOptionButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpBottomRight);
        }

        private void bottomRightOptionButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void sameFolderButtonOption_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpSameFolder);
        }

        private void sameFolderButtonOption_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void subFolderButtonOption_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpSubFolder);
        }

        private void subFolderButtonOption_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void subFolderTextBox_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpSubFolderText);
        }

        private void subFolderTextBox_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void generateButton_MouseHover(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpGenerate);
        }

        private void generateButton_MouseLeave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void imageWidthNumeric_Enter(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpImageWidth);
        }

        private void imageWidthNumeric_Leave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void imageHeightNumeric_Enter(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpImageHeight);
        }

        private void imageHeightNumeric_Leave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void imageLeftOffsetNumeric_Enter(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpImageLeftOffset);
        }

        private void imageLeftOffsetNumeric_Leave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void imageBottomOffsetNumeric_Enter(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpImageBottomOffset);
        }

        private void imageBottomOffsetNumeric_Leave(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }

        private void pageNumberNumeric_Enter(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = string.Format(Strings.HelpPageNumber);
        }

        private void pageNumberNumeric_FontChanged(object sender, EventArgs e)
        {
            this.toolStripStatus.Text = this.CurrentStatus;
        }
    }
}
