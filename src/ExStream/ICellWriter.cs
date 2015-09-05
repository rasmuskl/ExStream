using DocumentFormat.OpenXml;

namespace ExStream
{
    public interface ICellWriter
    {
        void WriteCell(OpenXmlWriter writer, object value);
    }
}