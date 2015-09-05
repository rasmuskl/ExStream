using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExStream.CellWriters
{
    public class StringCellWriter : ICellWriter
    {
        static readonly OpenXmlAttribute[] CellAttributes = { new OpenXmlAttribute("t", null, "inlineStr"), };

        public void WriteCell(OpenXmlWriter writer, object value)
        {
            writer.WriteStartElement(new Cell(), CellAttributes);

            writer.WriteStartElement(new InlineString());
            writer.WriteStartElement(new Text());
            writer.WriteString(value.ToString());
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}