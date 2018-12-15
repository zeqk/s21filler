using CommandLine;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using S21Filler.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace S21Filler
{
    class Program
    {

        class Options
        {

            [Value(0, MetaName = "s21template", Required = true, HelpText = "Plantilla PDF S-21")]
            public string S21Template { get; set; }

            [Value(1, MetaName = "excel", Required = true, HelpText = "Excel")]
            public string Excel { get; set; }

            [Value(2, MetaName = "year", Required = true, HelpText = "Año")]
            public int Year { get; set; }

            [Value(3, MetaName = "output-folder", Required = true, HelpText = "Carpeta destino")]
            public string OutputFolder { get; set; }


            
        }

        static void Main(string[] args)
        {

            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed(opts => RunOptionsAndReturnExitCode(opts))
                .WithNotParsed((errs) => HandleParseError(errs));

            
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            
        }

        private static void RunOptionsAndReturnExitCode(Options opts)
        {
            var records = getYearRecords(opts.Year, opts.Excel);

            //saveFileByPublisher(records, opts.S21Template, opts.OutputFolder);

            saveFileByFourPublishers(records, opts.S21Template, opts.OutputFolder);


        }

        static void saveFileByFourPublishers(IList<YearRecord> records, string template, string outputFolder)
        {
            var i = 0;
            foreach (var record in records)
            {
                using (var pdfReader = new PdfReader(template))
                {
                    var fourRecords = records.Skip(i).Take(4);

                    if (fourRecords.Any())
                    {

                        var output = Path.Combine(outputFolder, "S-21 - " + string.Join("; ", fourRecords.Select(r => r.Name)) + ".pdf");

                        using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(output, FileMode.Create)))
                        {

                            var pdfFormFields = pdfStamper.AcroFields;



                            for (int x = 1; x <= fourRecords.Count(); x++)
                            {
                                pdfFormFields.SetYearRecord(x, fourRecords.ElementAt(x - 1));
                                Console.WriteLine("S-21 generated for {0}", fourRecords.ElementAt(x - 1).Name);
                            }


                            pdfStamper.Close();
                        }
                    }
                }

                i = i + 4;
            }

            //PR Totals

            using (var pdfReader = new PdfReader(template))
            {
                var output = Path.Combine(outputFolder, "S-21 - TOTALES.pdf");

                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(output, FileMode.Create)))
                {

                    var pdfFormFields = pdfStamper.AcroFields;

                    var x = 1;
                    foreach (var type in (PublisherTypes[])Enum.GetValues(typeof(PublisherTypes)))
                    {
                        var totalYearRecord = getTotals(records, type);

                        pdfFormFields.SetYearRecord(x, totalYearRecord);
                        x++;
                        Console.WriteLine("S-21 generated for {0}", totalYearRecord.Name);
                    }
                    


                    pdfStamper.Close();
                }
            }

            
        }


        static void saveFileByPublisher(IList<YearRecord> records, string template, string outputFolder)
        {
            foreach (var record in records)
            {
                using (var pdfReader = new PdfReader(template))
                {
                    var output = Path.Combine(outputFolder, "S-21 - " + record.Name + ".pdf");

                    using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(output, FileMode.Create)))
                    {

                        var pdfFormFields = pdfStamper.AcroFields;
                        //pdfFormFields.GenerateAppearances = true;
                        pdfFormFields.SetYearRecord(1, record);

                        // flatten the form to remove editting options, set it to false

                        // to leave the form open to subsequent manual edits

                        //pdfStamper.FormFlattening = false;
                        //pdfStamper.AnnotationFlattening = false;                        

                        // close the pdf

                        pdfStamper.Close();
                    }
                }

                Console.WriteLine("S-21 generated for {0} {1} months", record.Name, record.Reports.Count);
            }

            //PR Totals
            foreach (var type in (PublisherTypes[])Enum.GetValues(typeof(PublisherTypes)))
            {
                var totalYearRecord = getTotals(records, type);

                using (var pdfReader = new PdfReader(template))
                {
                    var output = Path.Combine(outputFolder, "S-21 - " + totalYearRecord.Name + ".pdf");

                    using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(output, FileMode.Create)))
                    {

                        var pdfFormFields = pdfStamper.AcroFields;
                        //pdfFormFields.GenerateAppearances = true;
                        pdfFormFields.SetYearRecord(1, totalYearRecord);

                        // flatten the form to remove editting options, set it to false

                        // to leave the form open to subsequent manual edits

                        //pdfStamper.FormFlattening = false;
                        //pdfStamper.AnnotationFlattening = false;                        

                        // close the pdf

                        pdfStamper.Close();
                    }
                }
                Console.WriteLine("S-21 generated for {0}", totalYearRecord.Name);
            }
        }


        static YearRecord getTotals(IList<YearRecord> records, PublisherTypes type)
        {
            var rv = new YearRecord()
            {
                Anointed = string.Empty,
                DateOfBirth = null,
                E = false,
                Gender = null,
                HomeAddress = string.Empty,
                HomeTelephone = string.Empty,
                ImmersedDate = null,
                MobileTelephone = string.Empty,
                MS = false,
                Number = null,
                RP = false,
                Year = records.FirstOrDefault().Year
            
            };

            switch (type)
            {
                case PublisherTypes.RegularPioneer:
                    rv.Name = "PRECURSORES REGULARES";
                    break;
                case PublisherTypes.AuxiliarPioneer:
                    rv.Name = "PRECURSORES AUXILIARES";
                    break;
                case PublisherTypes.Publisher:
                default:
                    rv.Name = "PUBLICADORES";
                    break;
            }

            var monthsCount = records.Max(r => r.Reports.Count);

            //del 1 al 12
            for (int month = 1; month <= monthsCount; month++)
            {
                var monthReport = new MonthReport
                {
                    Month = month,
                    Placements = records.Sum(yr => yr.Reports.Where(mr => mr.Month == month && mr.Type == type).Sum(mr => mr.Placements)),
                    VideoShowings = records.Sum(yr => yr.Reports.Where(mr => mr.Month == month && mr.Type == type).Sum(mr => mr.VideoShowings)),
                    Hours = records.Sum(yr => yr.Reports.Where(mr => mr.Month == month && mr.Type == type).Sum(mr => mr.Hours)),
                    ReturnVisits = records.Sum(yr => yr.Reports.Where(mr => mr.Month == month && mr.Type == type).Sum(mr => mr.ReturnVisits)),
                    Studies = records.Sum(yr => yr.Reports.Where(mr => mr.Month == month && mr.Type == type).Sum(mr => mr.Studies)),
                    Remarks = records.SelectMany(yr => yr.Reports).Count(mr => mr.Month == month && mr.Type == type).ToString()
                };
                
                rv.Reports.Add(monthReport);
            }

            rv.Totals = new MonthReport
            {
                Placements = rv.Reports.Sum(r => r.Placements),
                VideoShowings = rv.Reports.Sum(r => r.VideoShowings),
                Hours = rv.Reports.Sum(r => r.Hours),
                ReturnVisits = rv.Reports.Sum(r => r.ReturnVisits)
            };


            return rv;
        }

        /// <summary>
        /// Extraction
        /// </summary>
        /// <param name="year"></param>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        static IList<YearRecord> getYearRecords(int year, string excelFile)
        {
            var rv = new List<YearRecord>();

            //FileInfo file = new FileInfo(@"C:\Users\zeqk\Desktop\TEMP\Actividad.xlsx");
            var file = new FileInfo(excelFile);
            try
            {
                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == year.ToString());

                    var i = 3;
                    YearRecord record;
                    do
                    {
                        i++;
                        record = getYearRecord(worksheet, i);
                        if (record != null)
                        {
                            record.Year = year;
                            rv.Add(record);
                        }

                    } while (record != null);

                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return rv;
        }

        static YearRecord getYearRecord(ExcelWorksheet worksheet, int row)
        {
            YearRecord rv = null;

            //FileInfo file = new FileInfo(@"C:\Users\zeqk\Desktop\TEMP\Actividad.xlsx");
            try
            {
                if ((string)worksheet.Cells[row, 1].Value != "TOTAL")
                {
                    rv = new YearRecord();

                    rv.Name = (string)worksheet.Cells[row, 1].Value;
                    rv.RP = (string)worksheet.Cells[row, 2].Value == "PR";
                    rv.HomeAddress = (string)worksheet.Cells[row, 117].Value;
                    rv.MobileTelephone = worksheet.Cells[row, 118].GetValue<string>();
                    rv.HomeTelephone = worksheet.Cells[row, 119].GetValue<string>();
                    rv.Gender = (string)worksheet.Cells[row, 120].Value == "F" ? Genders.Female : Genders.Male;
                    rv.DateOfBirth = worksheet.Cells[row, 121].GetValue<DateTime>();
                    rv.ImmersedDate = worksheet.Cells[row, 122].GetValue<DateTime>();
                    rv.Anointed = (string)worksheet.Cells[row, 123].Value;
                    rv.E = (string)worksheet.Cells[row, 124].Value == "SI";
                    rv.MS = (string)worksheet.Cells[row, 125].Value == "SI";
                    rv.RP = (string)worksheet.Cells[row, 126].Value == "SI";

                    for (int i = 1; i < 12; i++)
                    {
                        var columnIndex = 1; //Primera columna (Nombre)
                        if (i > 1)
                            columnIndex = (9 * (i -1)) + 1;
                        columnIndex++;
                        
                        if (worksheet.Cells[row, columnIndex + 3] != null && worksheet.Cells[row, columnIndex + 3].Value != null)
                        {
                            var type = PublisherTypes.Publisher;
                            switch (worksheet.Cells[row, columnIndex + 0].GetValue<string>())
                            {
                                case "PR":
                                    type = PublisherTypes.RegularPioneer;
                                    break;
                                case "PA":
                                    type = PublisherTypes.AuxiliarPioneer;
                                    break;
                                case "PUB":
                                default:
                                    type = PublisherTypes.Publisher;
                                    break;
                            }

                            rv.Reports.Add(new MonthReport
                            {
                                Type = type,
                                Placements = worksheet.Cells[row, columnIndex + 1].GetValue<int>(),
                                VideoShowings = worksheet.Cells[row, columnIndex + 2].GetValue<int>(),
                                Hours = worksheet.Cells[row, columnIndex + 3].GetValue<int>(),
                                ReturnVisits = worksheet.Cells[row, columnIndex + 4].GetValue<int>(),
                                Studies = worksheet.Cells[row, columnIndex + 5].GetValue<int>(),
                                Month = i,
                                Remarks = (type == PublisherTypes.AuxiliarPioneer ? "PA" : string.Empty) + " " + (worksheet.Cells[row, columnIndex + 1].Comment != null ? worksheet.Cells[row, columnIndex + 1].Comment.Text : string.Empty)
                            });
                        }
                        else
                            break;
                    }

                    rv.Totals = new MonthReport
                    {
                        Placements = rv.Reports.Sum(r => r.Placements),
                        VideoShowings = rv.Reports.Sum(r => r.VideoShowings),
                        Hours = rv.Reports.Sum(r => r.Hours),
                        ReturnVisits = rv.Reports.Sum(r => r.ReturnVisits)
                    };


                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            

            return rv;
        }


    }


    

    public static class AcroFieldsExtensions
    {
        public static void SetYearRecord(this AcroFields fields, int formNumber, YearRecord yearRecord)
        {
            var initIndex = getInitIndex(formNumber);

            if (yearRecord.Gender.HasValue)
            {
                fields.SetField("Check Box" + (initIndex + 1), yearRecord.Gender == Genders.Male ? "Yes" : "0");
                fields.SetField("Check Box" + (initIndex + 2), yearRecord.Gender == Genders.Female ? "Yes" : "0");
            }
            fields.SetFieldProperty("Text" + (initIndex + 3), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 3), yearRecord.Name + (!string.IsNullOrEmpty(yearRecord.Number) ? $" (#{yearRecord.Number})" : string.Empty));
            fields.SetFieldProperty("Text" + (initIndex + 4), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 4), yearRecord.HomeAddress);
            fields.SetFieldProperty("Text" + (initIndex + 5), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 5), yearRecord.HomeTelephone);
            fields.SetFieldProperty("Text" + (initIndex + 6), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 6), yearRecord.MobileTelephone);
            fields.SetFieldProperty("Text" + (initIndex + 7), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 7), yearRecord.DateOfBirth?.ToString("yyyy-MM-dd"));
            fields.SetFieldProperty("Text" + (initIndex + 8), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 8), yearRecord.ImmersedDate?.ToString("yyyy-MM-dd"));
            fields.SetFieldProperty("Text" + (initIndex + 9), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 9), yearRecord.Anointed);
            fields.SetField("Check Box" + (initIndex + 10), yearRecord.E ? "Yes" : "0");
            fields.SetField("Check Box" + (initIndex + 11), yearRecord.MS ? "Yes" : "0");
            fields.SetField("Check Box" + (initIndex + 12), yearRecord.RP ? "Yes" : "0");
            fields.SetFieldProperty("Text" + (initIndex + 12) + ".01", "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 12) + ".01", yearRecord.Year.ToString());

            foreach (var report in yearRecord.Reports.OrderBy(r => r.Month))
            {
                fields.SetMonthReport(formNumber, report);
            }

            fields.SetTotalReport(formNumber, yearRecord.Totals);

        }


        public static void SetTotalReport(this AcroFields fields, int formNumber, MonthReport report)
        {
            var initIndex = getInitIndexForMonth(formNumber, 13);

            fields.SetFieldProperty("Text" + (initIndex + 1), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 1), report.Placements.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 2), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 2), report.VideoShowings.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 3), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 3), report.Hours.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 4), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 4), report.ReturnVisits.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 6), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 6), report.Remarks);

        }

        public static void SetMonthReport(this AcroFields fields, int formNumber, MonthReport report)
        {
            var initIndex = getInitIndexForMonth(formNumber, report.Month);

            fields.SetFieldProperty("Text" + (initIndex + 1), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 1), report.Placements.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 2), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 2), report.VideoShowings.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 3), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 3), report.Hours.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 4), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 4), report.ReturnVisits.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 5), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 5), report.Studies.ToString());
            fields.SetFieldProperty("Text" + (initIndex + 6), "textsize", 8f, null);
            fields.SetField("Text" + (initIndex + 6), report.Remarks);

        }

        static int getInitIndexForMonth(int formNumber, int month)
        {
            var initIndex = getInitIndex(formNumber);

            initIndex += 12;

            for (int i = 1; i <= month; i++)
            {
                if(i > 1)
                {
                    initIndex += 6;
                }
            }

            return initIndex;

        }

        static int getInitIndex(int formNumber)
        {
            var initIndex = 0;

            switch (formNumber)
            {
                case 1:
                default:
                    initIndex += 0;
                    break;
                case 2:
                    initIndex += 90;
                    break;
                case 3:
                    initIndex += 180;
                    break;
                case 4:
                    initIndex += 270;
                    break;

            }
            return initIndex;

        }
    }
}
