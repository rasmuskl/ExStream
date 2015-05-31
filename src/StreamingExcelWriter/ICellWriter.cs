using DocumentFormat.OpenXml;

namespace StreamingExcelWriter
{
    public interface ICellWriter
    {
        void WriteCell(OpenXmlWriter writer, object value);
    }
}