using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using NUnit.Framework;
using OfficeOpenXml;

namespace ExStream.Tests.Playground
{
    public class SimplePerformanceTests
    {
        const int RowCount = 10000;
        const int ColumnCount = 100;
        const bool CheckMemory = true;
        const int MemoryCheckInterval = 5000;

        long _maxMemory;

        [Test]
        public void ExStreamWriterTest()
        {
            SimpleBenchmark(() =>
            {
                using (var writer = new ExStreamWriter("my-file-raw-test.xlsx"))
                {
                    using (var sheet = writer.WriteSheet("My Sheet"))
                    {
                        for (var i = 0; i < RowCount; i++)
                        {
                            for (var j = 0; j < ColumnCount; j++)
                            {
                                sheet.WriteCell("hopsa");
                            }

                            sheet.EndRow();
                            RegisterMaxMemory(i);
                        }

                        RegisterMaxMemory();
                    }

                    RegisterMaxMemory();
                }
            });
        }

        [Test]
        public void RawXmlTest()
        {
            SimpleBenchmark(() =>
            {
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

                                for (var r = 1; r <= RowCount; r++)
                                {
                                    writer.WriteStartElement("x", "row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                                    writer.WriteAttributeString("r", r.ToString());

                                    for (var c = 0; c < ColumnCount; c++)
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

                                    RegisterMaxMemory(r);
                                }

                                writer.WriteEndElement();
                                writer.WriteEndElement();
                            }

                            RegisterMaxMemory();
                        }

                        RegisterMaxMemory();
                    }

                    RegisterMaxMemory();
                }
            });
        }

        [Test]
        public void EPPlusTest()
        {
            SimpleBenchmark(() =>
            {

                var excelPackage = new ExcelPackage();

                var worksheet = excelPackage.Workbook.Worksheets.Add("My sheet");

                for (var i = 0; i < RowCount; i++)
                {
                    for (var j = 0; j < ColumnCount; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = "hopsa";
                    }

                    RegisterMaxMemory(i);
                }

                RegisterMaxMemory();
                excelPackage.SaveAs(new FileInfo("epplus-file.xlsx"));
            });
        }

        void SimpleBenchmark(Action action)
        {
            Console.WriteLine("Total memory before: " + GC.GetTotalMemory(true));
            var startNew = Stopwatch.StartNew();
            ResetMaxMemory();

            action();

            Console.WriteLine("Done in " + startNew.ElapsedMilliseconds + " ms");
            Console.WriteLine("Total memory after: " + GC.GetTotalMemory(true));
            ReportMaxMemory();
        }


        void ResetMaxMemory()
        {
            if (CheckMemory)
            {
                _maxMemory = GC.GetTotalMemory(true);
            }
        }

        void RegisterMaxMemory(int row = 0)
        {
            if (CheckMemory && row % MemoryCheckInterval == 0)
            {
                _maxMemory = Math.Max(_maxMemory, GC.GetTotalMemory(true));
            }
        }

        void ReportMaxMemory()
        {
            if (CheckMemory)
            {
                Console.WriteLine("Max memory: " + _maxMemory);
            }
        }
    }
}
