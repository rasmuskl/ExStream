using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExStream
{
    public class ExStreamWriter : IDisposable
    {
        readonly SpreadsheetDocument _document;
        readonly ExStreamWriterConfig _config;

        uint _nextSheetId = 1;
        uint _nextNumberFormatId = 165;

        public ExStreamWriter(string xlsxFile, ExStreamWriterConfig config = null)
        {
            _config = config ?? ExStreamWriterConfig.Current;
            _document = SpreadsheetDocument.Create(xlsxFile, SpreadsheetDocumentType.Workbook);
            InitialDocumentStructure();
        }

        public ExStreamWriter(Stream stream, ExStreamWriterConfig config = null)
        {
            _config = config ?? ExStreamWriterConfig.Current;
            _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            InitialDocumentStructure();
        }

        void InitialDocumentStructure()
        {
            _document.AddWorkbookPart();
            _document.WorkbookPart.Workbook = new Workbook();

            _document.WorkbookPart.Workbook.AppendChild(new Sheets());

            AddStyleSheet();
        }

        void AddStyleSheet()
        {
            WriteDefaultStyles();

            var stylePart = _document.WorkbookPart.WorkbookStylesPart;

            stylePart.Stylesheet.NumberingFormats = new NumberingFormats();

            var format = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(_nextNumberFormatId),
                FormatCode = StringValue.FromString(@"yyyy-mm-dd\ hh:mm;@")
            };

            stylePart.Stylesheet.NumberingFormats.AppendChild(format);

            var cellFormat = new CellFormat
            {
                NumberFormatId = _nextNumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
            };

            stylePart.Stylesheet.CellFormats.AppendChild(cellFormat);

            _nextNumberFormatId += 1;
        }

        void WriteDefaultStyles()
        {
            var stylePart = _document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            stylePart.Stylesheet = new Stylesheet();

            stylePart.Stylesheet.CellFormats = new CellFormats();

            var cellFormatZero = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
            };

            stylePart.Stylesheet.CellFormats.AppendChild(cellFormatZero);

            stylePart.Stylesheet.CellStyleFormats = new CellStyleFormats();

            var cellStyleFormatZero = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
            };

            stylePart.Stylesheet.CellStyleFormats.AppendChild(cellStyleFormatZero);

            stylePart.Stylesheet.Fonts = new Fonts();
            var fonts = stylePart.Stylesheet.Fonts;
            var font = new Font();
            font.AppendChild(new FontSize { Val = 11 });
            font.AppendChild(new Color { Theme = 1 });
            font.AppendChild(new FontName { Val = "Calibri" });
            font.AppendChild(new FontScheme { Val = FontSchemeValues.Minor });
            fonts.AppendChild(font);

            stylePart.Stylesheet.Fills = new Fills();
            stylePart.Stylesheet.Fills.AppendChild(new Fill(new PatternFill
            {
                PatternType = PatternValues.None,
            }));

            stylePart.Stylesheet.Borders = new Borders();
            var border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            stylePart.Stylesheet.Borders.AppendChild(border);
        }

        public ExStreamSheetWriter WriteSheet(string sheetName)
        {
            var worksheetPart = _document.WorkbookPart.AddNewPart<WorksheetPart>();

            var sheet = new Sheet()
            {
                Id = _document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = _nextSheetId,
                Name = sheetName
            };

            _nextSheetId += 1;
            _document.WorkbookPart.Workbook.Sheets.AppendChild(sheet);

            return new ExStreamSheetWriter(_config, worksheetPart);
        }

        public void Dispose()
        {
            _document?.Dispose();
        }
    }
}