using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace editdocuments
{
    public class TextArg : EventArgs
    {
        public TextArg(string text)
        {
            this.Text = text;
        }
        public string Text { get; set; }
    }

    public class CounterArgs : EventArgs
    {
        public CounterArgs(int value, int total)
        {
            this.Value = value;
            this.Total = total;
        }

        public int Total { get; set; }
        public int Value { get; set; }
    }

    public class IntArg : EventArgs
    {
        public IntArg(int value)
        {
            this.Value = value;
        }
        public int Value { get; set; }
    }
}
