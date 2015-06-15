using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using NUnit.Framework;

namespace ExStreamWriter.Tests
{
    public class WriterPerformanceVsRawTests
    {
        [Test]
        public void Writer()
        {
            var startNew = Stopwatch.StartNew();

            using (var writer = new ExcelWriter("my-file-raw-test.xlsx"))
            {
                Console.WriteLine("Total memory: " + GC.GetTotalMemory(true));

                using (var sheet = writer.WriteSheet("My Sheet"))
                {
                    for (var i = 0; i < 10000; i++)
                    {
                        for (var j = 0; j < 100; j++)
                        {
                            sheet.WriteCell("hopsa");
                        }

                        sheet.EndRow();
                    }
                }
            }

            Console.WriteLine("Done in " + startNew.ElapsedMilliseconds + " ms");
            Console.WriteLine("Total memory: " + GC.GetTotalMemory(true));
        }

        [Test]
        public void RawXml()
        {
            var startNew = Stopwatch.StartNew();

            if (File.Exists("test.xml"))
            {
                File.Delete("test.xml");
            }

            using (var fileStream = File.OpenWrite("test.zip"))
            {
                using (var zipArchive = new ZipArchive(fileStream, ZipArchiveMode.Create))
                {
                    var entry = zipArchive.CreateEntry("test.xml");

                    using (var stream = entry.Open())
                    {
                        using (var writer = new XmlTextWriter(stream, Encoding.UTF8))
                        {
                            writer.WriteStartElement("x", "worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                            writer.WriteStartElement("x", "sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                            for (var r = 1; r <= 10000; r++)
                            {
                                writer.WriteStartElement("x", "row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                                writer.WriteAttributeString("r", r.ToString());

                                for (var c = 0; c < 100; c++)
                                {
                                    writer.WriteStartElement("x", "c", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                                    writer.WriteAttributeString("t", "inlineStr");

                                    writer.WriteStartElement("x", "is", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                                    writer.WriteStartElement("x", "t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                                    writer.WriteString("hopsa");

                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                }

                                writer.WriteEndElement();

                            }

                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                    }

                }
            }

            Console.WriteLine("Done in " + startNew.ElapsedMilliseconds + " ms");
            Console.WriteLine("Total memory: " + GC.GetTotalMemory(true));
        }
    }
}
