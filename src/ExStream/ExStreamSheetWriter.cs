using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExStream
{
    public class ExStreamSheetWriter : IDisposable
    {
        readonly ExStreamWriterConfig _config;
        readonly OpenXmlWriter _writer;
        uint _nextRowNumber = 1;
        bool _inRow = false;

        public ExStreamSheetWriter(ExStreamWriterConfig config, WorksheetPart worksheetPart)
        {
            _config = config;
            _writer = OpenXmlWriter.Create(worksheetPart);

            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());
        }

        public void WriteRow(params object[] values)
        {
            if (_inRow)
            {
                EndRow();
            }

            EnsureRowStarted();

            foreach (var value in values)
            {
                WriteCell(value);
            }

            EndRow();
        }

        void EnsureRowStarted()
        {
            if (!_inRow)
            {
                _writer.WriteStartElement(new Row(), new[] { new OpenXmlAttribute("r", null, _nextRowNumber.ToString()), });
                _nextRowNumber += 1;
                _inRow = true;
            }
        }

        public void WriteCell(object value)
        {
            EnsureRowStarted();

            var cellWriter = _config.GetCellWriter(value.GetType());

            if (cellWriter == null)
            {
                throw new Exception("Unable to write cell with value type: {value.GetType()}");
            }

            cellWriter.WriteCell(_writer, value);
        }

        public void EndRow()
        {
            EnsureRowStarted();

            _writer.WriteEndElement();
            _inRow = false;
        }

        public void Dispose()
        {
            _writer?.WriteEndElement();
            _writer?.WriteEndElement();
            _writer?.Dispose();
        }
    }
}