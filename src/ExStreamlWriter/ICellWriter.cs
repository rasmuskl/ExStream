using DocumentFormat.OpenXml;

namespace ExStreamWriter
{
    public interface ICellWriter
    {
        void WriteCell(OpenXmlWriter writer, object value);
    }
}