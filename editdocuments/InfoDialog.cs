using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace editdocuments
{
    public partial class InfoDialog : Form
    {

        public InfoDialog(string title, string text, bool isRTF= true)
        {
            InitializeComponent();
            this.richTextBox1.ReadOnly = false; // require to load image
            SetTitle(title);
            if (isRTF)
            {
                LoadRTF(text);
            }
            else
            {
                LoadText(text);
            }
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.BackColor = Color.FromKnownColor(KnownColor.White);
        }

        public InfoDialog()
        {
            InitializeComponent();
            this.richTextBox1.BackColor = Color.FromKnownColor(KnownColor.White);
        }

        public void LoadTextFile(string textfile)
        {
            this.richTextBox1.LoadFile(textfile);
        }

        public void LoadText(string text)
        {
            this.richTextBox1.Text = text;
        }

        public void LoadRTF(string text)
        {
            this.richTextBox1.Rtf = text;
        }

        public void SetTitle(string title)
        {
            this.Text = title;
        }
    }
}
