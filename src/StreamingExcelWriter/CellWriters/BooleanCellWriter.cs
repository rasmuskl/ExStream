using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace StreamingExcelWriter.CellWriters
{
    public class BooleanCellWriter : ICellWriter
    {
        public void WriteCell(OpenXmlWriter writer, object value)
        {
            writer.WriteStartElement(new Cell(), new [] { new OpenXmlAttribute("t", null, "b"),  });

            var b = (bool) value;

            writer.WriteStartElement(new CellValue());
            writer.WriteString(BooleanValue.FromBoolean(b).ToString());
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}