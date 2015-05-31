using System;
using System.IO;
using NUnit.Framework;

namespace StreamingExcelWriter.Tests
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public void CanWriteRows()
        {
            using (var writer = new ExcelWriter("my-file-rows.xlsx"))
            {
                using (var sheet = writer.WriteSheet("MySheet"))
                {
                    for (var i = 1; i < 1000; i++)
                    {
                        sheet.WriteRow("1", 2, Math.PI, true, false, DateTime.Now, Guid.NewGuid());
                    }
                }
            }
        }

        [Test]
        public void CanWriteCells()
        {
            using (var writer = new ExcelWriter("my-file-cells.xlsx"))
            {
                using (var sheet = writer.WriteSheet("MySheet"))
                {
                    for (var i = 0; i < 1000; i++)
                    {
                        for (var j = 0; j < 100; j++)
                        {
                            sheet.WriteCell("hopsa");
                        }

                        sheet.EndRow();
                    }
                }
            }
        }

        [Test]
        public void CanWriteToMemoryStream()
        {
            using (var stream = new MemoryStream())
            {
                using (var writer = new ExcelWriter(stream))
                {
                    using (var sheet = writer.WriteSheet("MySheet"))
                    {
                        for (var i = 1; i < 1000; i++)
                        {
                            sheet.WriteRow("1", 2, 3);
                        }
                    }
                }

                stream.Seek(0, SeekOrigin.Begin);

                File.WriteAllBytes("my-streamed-file.xlsx", stream.ToArray());
            }
        }
    }
}