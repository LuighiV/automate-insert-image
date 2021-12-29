using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace editdocuments
{
    // From https://stackoverflow.com/a/18727100 
    public class ControlWriter : TextWriter
    {
        private TextBox textbox;
        public ControlWriter(TextBox textbox)
        {
            this.textbox = textbox;
        }

        public override void Write(char value)
        {
            if (textbox.InvokeRequired)
            {
                try
                {
                    textbox.Invoke(new MethodInvoker(() => textbox.Text += value));
                }
                catch (ObjectDisposedException ex)
                {
                    Console.WriteLine(" (" + ex.Message + "): " + ex);
                }
            }
            else
            {
                textbox.Text += value;
            }

            textbox.SelectionStart = textbox.TextLength;
            textbox.ScrollToCaret();
        }

        public override void Write(string value)
        {
            if (textbox.InvokeRequired)
            {
                try
                {
                    textbox.Invoke(new MethodInvoker(() => textbox.AppendText(value.ToString())));
                }
                catch (ObjectDisposedException ex)
                {
                    Console.WriteLine(" (" + ex.Message + "): " + ex);
                }
            }
            else
            {
                textbox.AppendText(value.ToString());
            }
        }

        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }
    }
}
