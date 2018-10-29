using CommandLine;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using S21Filler.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

            //string pdfTemplate = @"C:\Users\zeqk\Desktop\TEMP\S-21-S_4up.pdf";

            
            foreach (var record in records)
            {
                using (var pdfReader = new PdfReader(opts.S21Template))
                {
                    var output = Path.Combine(opts.OutputFolder, "S-21 - " + record.Name + ".pdf");

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
            }
                
            

        }


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
                            rv.Add(record);

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
                        var init = 1;
                        if (i > 1)
                            init = 10 * i;
                        init++;
                        if (worksheet.Cells[row, init + 3] != null && worksheet.Cells[row, init + 3].Value != null)
                        {
                            rv.Reports.Add(new MonthReport
                            {
                                Placements = worksheet.Cells[row, init + 1].GetValue<int>(),
                                VideoShowings = worksheet.Cells[row, init + 2].GetValue<int>(),
                                Hours = worksheet.Cells[row, init + 3].GetValue<int>(),
                                ReturnVisits = worksheet.Cells[row, init + 4].GetValue<int>(),
                                Studies = worksheet.Cells[row, init + 5].GetValue<int>(),
                                Month = i,
                                Remarks = worksheet.Cells[row, init + 1].Comment != null ? worksheet.Cells[row, init + 1].Comment.Text : null
                            });
                        }
                        else
                            break;
                    }
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

            fields.SetField("Check Box" + (initIndex + 1), yearRecord.Gender == Genders.Male ? "Yes" : "0");
            fields.SetField("Check Box" + (initIndex + 2), yearRecord.Gender == Genders.Female ? "Yes" : "0");
            fields.SetField("Text" + (initIndex + 3), yearRecord.Name + $" (#{yearRecord.Number})");
            fields.SetField("Text" + (initIndex + 4), yearRecord.HomeAddress);
            fields.SetField("Text" + (initIndex + 5), yearRecord.HomeTelephone);
            fields.SetField("Text" + (initIndex + 6), yearRecord.MobileTelephone);
            fields.SetField("Text" + (initIndex + 7), yearRecord.DateOfBirth.ToString("yyyy-MM-dd"));
            fields.SetField("Text" + (initIndex + 8), yearRecord.ImmersedDate.HasValue ? yearRecord.ImmersedDate.Value.ToString("yyyyy-MM-dd") : "");
            fields.SetField("Text" + (initIndex + 9), yearRecord.Anointed);
            fields.SetField("Check Box" + (initIndex + 10), yearRecord.E ? "Yes" : "0");
            fields.SetField("Check Box" + (initIndex + 11), yearRecord.MS ? "Yes" : "0");
            fields.SetField("Check Box" + (initIndex + 12), yearRecord.RP ? "Yes" : "0");
            fields.SetField("Text" + (initIndex + 12) + ".01", yearRecord.Year.ToString());

            foreach (var report in yearRecord.Reports.OrderBy(r => r.Month))
            {
                fields.SetMonthReport(formNumber, report);
            }

        }

        public static void SetMonthReport(this AcroFields fields, int formNumber, MonthReport report)
        {
            var initIndex = getInitIndexForMonth(formNumber, report.Month);

            fields.SetField("Text" + (initIndex + 1), report.Placements.ToString());
            fields.SetField("Text" + (initIndex + 2), report.VideoShowings.ToString());
            fields.SetField("Text" + (initIndex + 3), report.Hours.ToString());
            fields.SetField("Text" + (initIndex + 4), report.ReturnVisits.ToString());
            fields.SetField("Text" + (initIndex + 5), report.Studies.ToString());
            fields.SetField("Text" + (initIndex + 6), report.Remarks);

        }

        static int getInitIndexForMonth(int formNumber, int month)
        {
            var initIndex = getInitIndex(formNumber);

            initIndex += 12;

            for (int i = 1; i < month; i++)
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
