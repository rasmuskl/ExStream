using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExStreamWriter.CellWriters
{
    public class DateTimeCellWriter : ICellWriter
    {
        public void WriteCell(OpenXmlWriter writer, object value)
        {
            var dateTime = (DateTime) value;

            writer.WriteStartElement(new Cell(), new [] { new OpenXmlAttribute("s", null, "1"), });

            writer.WriteStartElement(new CellValue());
            writer.WriteString(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}