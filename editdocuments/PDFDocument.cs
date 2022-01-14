using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.IO.Image;
using iText.Layout;
using iText.Layout.Element;

namespace editdocuments
{
    internal class PDFDocument
    {
        public string DocumentPath { get; set; }

        public PdfDocument pdfDocument;
        public Document document;
        public TextLocationStrategy strategy;
        public PdfCanvasProcessor processor;

        public string inputPath { get; set; }
        public string outputPath {get; set;}
        public textChunk reference;
        public int pageNumber;
        public Coords referencePoint;

        public PDFDocument(string inputPath, string outputPath,
            bool Verbose = false)
        {
            this.inputPath = inputPath;
            this.outputPath = outputPath;
            this.pdfDocument = new PdfDocument(new PdfReader(inputPath),
                new PdfWriter(outputPath));
            this.document = new Document(pdfDocument);

            FilteredEventListener listener = new FilteredEventListener();
            this.strategy = listener.AttachEventListener(new TextLocationStrategy());
            this.processor = new PdfCanvasProcessor(listener);

            this.referencePoint = new Coords(0, 0);
        }


        internal void Close(bool Verbose = false)
        {

            this.document.Close();

            if (Verbose)
            {
                Console.WriteLine(Strings.InfoSavedPDFFile, this.outputPath);
            }

        }

        public void AddPicture(string picturePath, double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
             double Height = 150)
        {
            AddPictureInPage(picturePath, LeftOffset, BottomOffset, Width, Height, this.pageNumber);
        }

        internal void AddPictureInPage(string picturePath, double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
             double Height = 150, int page=1)
        {
            Coords customReference = (reference == null) ? 
                this.referencePoint:
                new Coords(reference.rect.GetX(), reference.rect.GetY());

            AddPictureInPageReference(picturePath,
                customReference, LeftOffset, BottomOffset, Width, Height, page);
        }


        internal void AddPictureInPageReference(string picturePath,
            double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
            double Height = 150, int page = 1)
        {
            AddPictureInPageReference(picturePath,
                this.referencePoint, LeftOffset, BottomOffset, Width, Height, page);
        }

        internal void AddPictureInPageReference(string picturePath, PageReference pageReference,
            double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
            double Height = 150, int page = 1)
        {

            Coords customReference = getPageReferencePoint(pageReference, page);
            AddPictureInPageReference(picturePath,
                customReference, LeftOffset, BottomOffset, Width, Height, page);
        }

        public void AddPictureInPageReference(string picturePath, Coords customReference, 
            double LeftOffset = 0, double BottomOffset = 0, double Width = 200,
            double Height = 150, int page = 1)
        {
            Coords position = new Coords(customReference.X + LeftOffset, customReference.Y + BottomOffset);

            ImageData imageData = ImageDataFactory.Create(picturePath);
            Image image = new Image(imageData).ScaleAbsolute((float)Width, (float)Height).
                SetFixedPosition(page, (float)position.X, (float)position.Y);
            this.document.Add(image);
        }

        public bool GetReference(string text, bool Verbose = false)
        {
            for (int i=0; i < this.pdfDocument.GetNumberOfPages(); i++)
            {
                int pageNumber = i + 1;
                var result = findTextInPage(text, pageNumber, Verbose);
                if(result != null)
                {
                    this.reference = result;
                    this.pageNumber = pageNumber;
                    return true;
                }
            }
            return false;
        }

        public bool HasPage(int page)
        {
            return (page > 0) & (page <= this.pdfDocument.GetNumberOfPages());
        }

        public textChunk findTextInPage(string text, int page, bool Verbose = false)
        {
            var chunks = GetTextChunksInPage(page, reorder: true, Verbose:Verbose);
            return findTextInTextChunks(text,chunks,Verbose);
        }

        public textChunk findTextInTextChunks(string findText, List<textChunk> chunks, 
            bool Verbose=false)
        {
            string fulltext = chunks.Aggregate("", (accum, element) => accum + element.text);
            if (Verbose)
            {
                Console.WriteLine(fulltext);
            }

            var found = fulltext.IndexOf(findText);
            if (Verbose)
            {
                Console.WriteLine(found);
            }

            if (found > 0)
            {
                return chunks[found];
            }
            return null;
        }

        public List<textChunk> GetTextChunksInPage(int page, bool reorder = true, bool Verbose=false)
        {
            this.strategy.ClearTextChunks();
            this.processor.ProcessPageContent(this.pdfDocument.GetPage(page));

            if (Verbose)
            {
                this.strategy.printTextChunks();
            }

            if (!reorder)
            {
                return this.strategy.GetTextChunks();
            }
            return ReOrderByPositionInPage(this.strategy.GetTextChunks());
        }

        public List<textChunk> ReOrderByPositionInPage(List<textChunk> chunks, bool verbose=false)
        {
            var groupedList = chunks.GroupBy(u => u.rect.GetY()).OrderByDescending(grp => grp.Key)
                .Select(grp => grp.OrderBy(element => element.rect.GetX()).ToList())
                .ToList();

            if (verbose)
            {
                foreach (List<textChunk> item in groupedList)
                {

                    string text = item.Aggregate("", (accum, element) => accum + element.text);
                    Console.WriteLine(text);
                }
            }

            return groupedList.SelectMany(x => x).ToList();
        }

        public Coords getPageReferencePoint(PageReference pageReference, int page)
        {
            return getPageReferencePoint(pageReference, page, pdfDocument);
        }

        static public Coords getPageReferencePoint(PageReference pageReference, int page, PdfDocument pdfdocument)
        {
            var pagesize = pdfdocument.GetPage(page).GetPageSize();
            switch (pageReference)
            {
                case PageReference.top_left:
                    return new Coords(0, pagesize.GetHeight());
                case PageReference.top_right:
                    return new Coords(pagesize.GetWidth(), pagesize.GetHeight());
                case PageReference.bottom_right:
                    return new Coords(pagesize.GetWidth(), 0);
                default:
                    return new Coords(0, 0);
            }
        }
    }

    /// <summary>
    /// Reference https://stackoverflow.com/a/43785366
    /// </summary>
    public class TextLocationStrategy : LocationTextExtractionStrategy
    {
        private List<textChunk> objectResult = new List<textChunk>();

        public override void EventOccurred(IEventData data, EventType type)
        {
            if (!type.Equals(EventType.RENDER_TEXT))
                return;

            TextRenderInfo renderInfo = (TextRenderInfo)data;

            string curFont = renderInfo.GetFont().GetFontProgram().ToString();

            float curFontSize = renderInfo.GetFontSize();

            IList<TextRenderInfo> text = renderInfo.GetCharacterRenderInfos();

            foreach (TextRenderInfo t in text)
            {
                string letter = t.GetText();
                Vector letterStart = t.GetBaseline().GetStartPoint();
                Vector letterEnd = t.GetAscentLine().GetEndPoint();
                Rectangle letterRect = new Rectangle(letterStart.Get(0), letterStart.Get(1), letterEnd.Get(0) - letterStart.Get(0), letterEnd.Get(1) - letterStart.Get(1));

                //if (letter != " " && !letter.Contains(' '))
                //{
                textChunk chunk = new textChunk
                {
                    text = letter,
                    rect = letterRect,
                    fontFamily = curFont,
                    fontSize = curFontSize,
                    spaceWidth = t.GetSingleSpaceWidth() / 2f
                };

                objectResult.Add(chunk);
                //}
            }
        }
        public void printTextChunks()
        {
            foreach (textChunk chunk in objectResult)
            {
                Console.WriteLine($"Text {chunk.text} found at ({chunk.rect.GetX()},{chunk.rect.GetY()})");
            }
        }
        public List<textChunk> GetTextChunks()
        {
            return objectResult;
        }

        public void ClearTextChunks()
        {
            objectResult.Clear();
        }
    }
    public class textChunk
    {
        public string text { get; set; }
        public Rectangle rect { get; set; }
        public string fontFamily { get; set; }
        public float fontSize { get; set; }
        public float spaceWidth { get; set; }
    }

    public struct Coords
    {
        public Coords (double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; set; }
        public double Y { get; set; }
    }
    public enum PageReference
    {
        bottom_left = 0,
        top_left,
        top_right,
        bottom_right
    }
}
