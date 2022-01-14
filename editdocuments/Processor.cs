using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace editdocuments
{
    public interface Processor
    {
        void RunProcess(DataInfo data, bool verbose=false);

        event EventHandler StartProcessing;
        event EventHandler<CounterArgs> UpdateCounter;
        event EventHandler<TextArg> StartProcessingFile;
        event EventHandler<TextArg> GotPlaceHolderPosition;
        event EventHandler<TextArg> PDFSaved;
        event EventHandler<TextArg> FinishProcessingFile;
        event EventHandler FinishProcessing;
    }
}
