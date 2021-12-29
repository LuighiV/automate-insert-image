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

namespace editdocuments
{
    public partial class GUI : Form
    {

        public bool SelectedFiles = true;
        public bool KeepAspect = true;
        public bool IgnoreTextChanged = false;
        public IEnumerable<string> SelectedPaths = null;
        public string FolderPath = null;
        public string ImagePath = null;
        public string PlaceHolder = null;
        public Image CurrentImage = null;

        public Quantity ImageWidth = Quantity.FromInches(0);
        public Quantity ImageHeight = Quantity.FromInches(0);
        public Quantity ImageLeftOffset = Quantity.FromInches(0);
        public Quantity ImageBottomOffset = Quantity.FromInches(0);

        public GUI()
        {
            InitializeComponent();
            this.comboBox1.DataSource = Globals.AvailableUnits;
            this.comboBox1.DisplayMember = "Literal";
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (SelectedFiles)
            {
                this.SelectedPaths = this.textBox1.Text.Split(',');
            }
            else
            {
                this.FolderPath = this.textBox1.Text;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.SelectedFiles)
            {
                
                this.openFileDialog1.ShowDialog();
            }
            else
            {
                DialogResult result =  this.folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(this.folderBrowserDialog1.SelectedPath))
                {
                    this.FolderPath = this.folderBrowserDialog1.SelectedPath;
                    this.textBox1.Text = this.FolderPath;
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.SelectedFiles = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            this.SelectedFiles = false;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            this.SelectedPaths = this.openFileDialog1.FileNames;
            this.textBox1.Text = String.Join(",", this.SelectedPaths);
            this.textBox1.Enabled = true;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.openFileDialog2.ShowDialog();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            this.IgnoreTextChanged = true;

            this.ImagePath = this.openFileDialog2.FileName;
            this.textBox2.Text = this.ImagePath;

            this.CurrentImage = new Bitmap(this.ImagePath);
            this.numericUpDown1.Value = this.CurrentImage.Width;
            this.numericUpDown2.Value = this.CurrentImage.Height;

            fillPictureBox(this.pictureBox1, this.CurrentImage);

            InitiateDimensions();
            ConvertUnits();
            UpdateDimensionsInUI();

            this.textBox2.Enabled = true;

            this.IgnoreTextChanged = false;
        }

        private void InitiateDimensions()
        {
            this.ImageWidth = Quantity.FromPixels((double)this.numericUpDown1.Value);
            this.ImageHeight = Quantity.FromPixels((double)this.numericUpDown2.Value);
            this.ImageLeftOffset = Quantity.FromPixels((double)this.numericUpDown3.Value);
            this.ImageBottomOffset = Quantity.FromPixels((double)this.numericUpDown4.Value);
        }

        private void ConvertUnits()
        {
            UnitStruct currentUnit = (UnitStruct) this.comboBox1.SelectedItem;
            this.ImageWidth.ToUnit(currentUnit.Unit);
            this.ImageHeight.ToUnit(currentUnit.Unit);
            this.ImageLeftOffset.ToUnit(currentUnit.Unit);
            this.ImageBottomOffset.ToUnit(currentUnit.Unit);
        }

        private void UpdateDimensionsInUI()
        {
            this.numericUpDown1.Value = (decimal)this.ImageWidth.Value;
            this.numericUpDown2.Value = (decimal)this.ImageHeight.Value;
            this.numericUpDown3.Value = (decimal)this.ImageLeftOffset.Value;
            this.numericUpDown4.Value = (decimal)this.ImageBottomOffset.Value;
        }

        private void quantityBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConvertUnits();
            UpdateDimensionsInUI();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            this.ImagePath = this.textBox2.Text;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.KeepAspect = this.checkBox1.Checked;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;

            double aspectRatio = this.ImageWidth.Value / this.ImageHeight.Value;

            this.ImageWidth.Value = (double) this.numericUpDown1.Value;
            if (this.KeepAspect)
            {
                this.IgnoreTextChanged = true;
                this.ImageHeight.Value = this.ImageWidth.Value / aspectRatio;
                this.numericUpDown2.Value = (decimal)this.ImageHeight.Value;
                this.IgnoreTextChanged = false;
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;

            double aspectRatio = this.ImageWidth.Value / this.ImageHeight.Value;

            this.ImageHeight.Value = (double)this.numericUpDown2.Value;
            if (this.KeepAspect)
            {
                this.IgnoreTextChanged = true;
                this.ImageWidth.Value = this.ImageHeight.Value * aspectRatio;
                this.numericUpDown1.Value = (decimal)this.ImageWidth.Value;
                this.IgnoreTextChanged = false;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            this.PlaceHolder = this.textBox3.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!ValidateChildren(ValidationConstraints.Enabled))
                return;

            try
            {
                this.Enabled = false;
                Application.UseWaitCursor = true;
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();
                var processForm = new ProcessDialog();
                processForm.Show();

                processForm.RunGUIProcess(this.ImagePath,
                    this.PlaceHolder,
                    this.ImageLeftOffset.ConvertValue(GUnits.point),
                    this.ImageBottomOffset.ConvertValue(GUnits.point),
                    this.ImageWidth.ConvertValue(GUnits.point),
                    this.ImageHeight.ConvertValue(GUnits.point),
                    this.SelectedFiles ? null : this.FolderPath,
                    this.SelectedFiles ? this.SelectedPaths : null,
                    false,
                    null,
                    "PDF",
                    false,
                    true);
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

                MessageBox.Show(Strings.ExecutionCompletedDefaultMessage,
                    Strings.ExecutionCompletedTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

            }

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;
            this.ImageLeftOffset.Value = (double)this.numericUpDown3.Value;
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            if (this.IgnoreTextChanged)
                return;
            this.ImageBottomOffset.Value = (double)this.numericUpDown4.Value;
        }

        // Funtions to verify data

        private bool IsValidSelectedPath()
        {
            return (!String.IsNullOrEmpty(textBox1.Text));
        }

        private bool IsValidPicturePath()
        {
            return (!String.IsNullOrEmpty(textBox2.Text));
        }

        private bool IsValidPlaceHolder()
        {
            return (!String.IsNullOrEmpty(textBox3.Text));
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
                float size_ratio = (float)pbox.Height / bmp.Height;
                int sample_width = (int)(pbox.Width / size_ratio);
                src_rect = new Rectangle((bmp.Width - sample_width) / 2, 0, sample_width, bmp.Height);
            }
            else
            {
                float size_ratio = (float)pbox.Width / bmp.Width;
                int sample_height = (int)(pbox.Height / size_ratio);
                src_rect = new Rectangle(0, (bmp.Height - sample_height) / 2, bmp.Width, sample_height);
            }

            g.DrawImage(bmp, dest_rect, src_rect, GraphicsUnit.Pixel);
            g.Dispose();

            pbox.Image = resized;
        }


        private void textBox1_Validating(object sender, CancelEventArgs e)
        {

            //Console.WriteLine(sender.GetType().Name);
            if (IsValidSelectedPath())
            {
                errorProvider1.SetError(this.textBox1, String.Empty);
                e.Cancel = false;
            }
            else
            {
                errorProvider1.SetError(this.textBox1, Strings.TextValidationFilePath);
                e.Cancel = this.button3.Focused;
            }
            
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
            if (IsValidPicturePath())
            {
                e.Cancel = false;
                errorProvider2.SetError(this.textBox2, String.Empty);
                
            }
            else
            {
                e.Cancel = this.button3.Focused;
                errorProvider2.SetError(this.textBox2, Strings.TextValidationPicturePath);
                
            }
        }

        private void textBox3_Validating(object sender, CancelEventArgs e)
        {
            if (IsValidPlaceHolder())
            {
                errorProvider3.SetError(this.textBox3, String.Empty);
                e.Cancel = false;
            }
            else
            {
                errorProvider3.SetError(this.textBox3, Strings.TextValidationPlaceholder);
                e.Cancel = this.button3.Focused;
            }
        }
    }
}
