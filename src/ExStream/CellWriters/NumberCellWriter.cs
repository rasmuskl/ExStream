using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExStream.CellWriters
{
    public class NumberCellWriter : ICellWriter
    {
        public void WriteCell(OpenXmlWriter writer, object value)
        {
            writer.WriteStartElement(new Cell());

            writer.WriteStartElement(new CellValue());
            writer.WriteString(string.Format(CultureInfo.InvariantCulture, "{0}", value));
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}