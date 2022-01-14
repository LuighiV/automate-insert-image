using System.Configuration;
using System.Collections.Generic;

namespace editdocuments
{
    public struct DataInfo
    {

        public DataInfo(
            bool fromSettings = true,
            string picturePath = null,
            string textPlaceHolder = null,
            double leftOffset = 0,
            double bottomOffset = 0,
            double width = 200,
            double height = 150,
            GUnits unit = 0,
            string inputPath = null,
            bool isFilesSelected = true,
            bool wordAppVisible = false,
            bool isSubFolderSelected = true,
            string folderSave = null,
            string subFolderSave = null,
            bool saveFile = false,
            bool isAbsolute = false,
            bool bottomLeftSelected = true,
            bool topLeftSelected = false,
            bool topRightSelected = false,
            bool bottomRightSelected = false,
            int pageNumber = 1,
            bool isWordDocumentType = true)
        {
            PicturePath = (fromSettings)? Properties.Settings.Default.PicturePath: picturePath;
            TextPlaceHolder = (fromSettings) ? Properties.Settings.Default.PlaceHolder : textPlaceHolder;
            
            InputPath = (fromSettings) ? Properties.Settings.Default.InputPath : inputPath;
            IsFilesSelected = (fromSettings) ? Properties.Settings.Default.IsFilesOptionSelected : isFilesSelected;
            WordAppVisible = wordAppVisible;
            IsSubFolderSelected = (fromSettings) ? Properties.Settings.Default.IsSubFolderSelected : isSubFolderSelected;
            _FolderSave = folderSave;
            _SubFolderSave = (fromSettings) ? Properties.Settings.Default.SubFolderSave : subFolderSave;
            SaveFile = saveFile;

            IsAbsolute = (fromSettings) ? !Properties.Settings.Default.IsRelativeToTextSelected : isAbsolute;
            
            PageNumber = (fromSettings) ? (int) Properties.Settings.Default.PageNumber:  pageNumber;

            _ImageWidth = new Quantity(0,unit);
            _ImageHeight = new Quantity(0, unit);
            _ImageLeftOffset = new Quantity(0, unit);
            _ImageBottomOffset = new Quantity(0, unit);
            _Unit = unit;

            Reference = PageReference.bottom_left;
            Type = DocumentType.Word;

            BottomLeftSelected = (fromSettings) ? Properties.Settings.Default.IsRelativeToTextSelected : bottomLeftSelected;
            TopLeftSelected = (fromSettings) ? Properties.Settings.Default.IsRelativeToTextSelected : topLeftSelected;
            TopRightSelected = (fromSettings) ? Properties.Settings.Default.IsRelativeToTextSelected : topRightSelected;
            BottomRightSelected = (fromSettings) ? Properties.Settings.Default.IsRelativeToTextSelected : bottomRightSelected;

            IsWordDocumentType = (fromSettings) ? Properties.Settings.Default.IsWordDocumentType : isWordDocumentType;

            LeftOffset = (fromSettings) ? (double)Properties.Settings.Default.ImageLeftOffset : leftOffset;
            BottomOffset = (fromSettings) ? (double)Properties.Settings.Default.ImageBottomOffset : bottomOffset;
            Width = (fromSettings) ? (double)Properties.Settings.Default.ImageWidth : width;
            Height = (fromSettings) ? (double)Properties.Settings.Default.ImageHeight : height;
        }

        private Quantity _ImageWidth;
        private Quantity _ImageHeight;
        private Quantity _ImageLeftOffset;
        private Quantity _ImageBottomOffset;
        private GUnits _Unit;
        private string _FolderSave;
        private string _SubFolderSave;


        public GUnits Unit { get { return _Unit; } set { _Unit = value; 
                _ImageWidth.ToUnit(value);
                _ImageHeight.ToUnit(value);
                _ImageLeftOffset.ToUnit(value);
                _ImageBottomOffset.ToUnit(value);
            } }

        public string PicturePath { get; set; }
        public string TextPlaceHolder { get; set; }
        public double LeftOffset { get => _ImageLeftOffset.Value; set => _ImageLeftOffset.Value = value; }
        public double BottomOffset { get => _ImageBottomOffset.Value; set => _ImageBottomOffset.Value = value; }
        public double Width { get => _ImageWidth.Value; set => _ImageWidth.Value = value; }
        public double Height { get => _ImageHeight.Value; set => _ImageHeight.Value = value; } 
        public string InputPath { get; set; }
        public bool IsFilesSelected { get; set; }
        public bool WordAppVisible { get; set; }
        public bool IsSubFolderSelected { get; set; }

        public bool SaveFile { get; set; }

        public bool IsAbsolute { get; set; }

        public PageReference Reference { get; set; }

        public int PageNumber { get; set; }

        public DocumentType Type { get; set; }

        public IEnumerable<string> InputFilePaths => (IsFilesSelected) ? InputPath.Split(',') : null;

        public string FolderPath => (IsFilesSelected) ? null : InputPath;

        public string FolderSave { get => (IsSubFolderSelected)? null:_FolderSave;
            set => _FolderSave = value; }
        public string SubFolderSave { get => (IsSubFolderSelected) ? _SubFolderSave:null; 
            set => _SubFolderSave = value; }


        public bool BottomLeftSelected
        {
            get { return Reference == PageReference.bottom_left; }
            set
            {
                if (value)
                {
                    Reference = PageReference.bottom_left;
                }
            }
        }

        public bool TopLeftSelected
        {
            get { return Reference == PageReference.top_left; ; }
            set
            {
                if (value)
                {
                    Reference = PageReference.top_left;
                }
            }
        }

        public bool TopRightSelected
        {
            get { return Reference == PageReference.top_right; }
            set
            {
                if (value)
                {
                    Reference = PageReference.top_right;
                }
            }
        }

        public bool BottomRightSelected
        {
            get { return Reference == PageReference.bottom_right; ; }
            set
            {
                if (value)
                {
                    Reference = PageReference.bottom_right;
                }
            }
        }

        public bool IsWordDocumentType
        {
            get { return Type == DocumentType.Word; ; }
            set
            {
                if (value) 
                {
                    Type = DocumentType.Word; 
                }
                else 
                { 
                    Type = DocumentType.PDF;
                }
            }
        }
    }
}