using iTextSharp.text.pdf;
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
        static void Main(string[] args)
        {
            string pdfTemplate = @"C:\Users\perez.e\Desktop\TEMP\S-21-S_4up.pdf";

            string newFile = @"C:\Users\perez.e\Desktop\TEMP\S-21-S_4up3.pdf";

            var record = new YearRecord()
            {
                Name = "Juan Jorge",
                Number = "321312",
                HomeAddress = "Callao 23, Claypole",
                HomeTelephone = "42335544",
                MobileTelephone = "1545487899",
                DateOfBirth = new DateTime(1988, 01, 01),
                ImmersedDate = new DateTime(1998, 01, 01),
                Anointed = "O.O.",
                Gender = Genders.Male,
                E = true,
                MS = false,
                RP = true,
                Year = 2018
            };
            record.Reports.Add(new MonthReport()
            {
                Month =1,
                Hours = 70,
                Placements = 10,
                Studies =1,
                ReturnVisits = 10,
                VideoShowings = 12,
                Remarks ="Enfermedad"
            });

            var pdfReader = new PdfReader(pdfTemplate);

            var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));

            var pdfFormFields = pdfStamper.AcroFields;
            pdfFormFields.SetYearRecord(1, record);

            //pdfFormFields.SetField("Check Box1", "Yes");
            //pdfFormFields.SetField("Check Box2", "0");
            //pdfFormFields.SetField("Text3", "");
            //pdfFormFields.SetField("Text4", "");
            //pdfFormFields.SetField("Text5", "42191111");
            //pdfFormFields.SetField("Text6", "1522465465");
            //pdfFormFields.SetField("Text7", "25/11/1988");
            //pdfFormFields.SetField("Text8", "13/4/2003");
            //pdfFormFields.SetField("Text9", "O.O.");
            //pdfFormFields.SetField("Check Box10", "Yes");
            //pdfFormFields.SetField("Check Box11", "0");
            //pdfFormFields.SetField("Check Box12", "Yes");
            //pdfFormFields.SetField("Text12.01", "2018");


            //for (int i = 13; i < 90; i++)
            //{

            //}

            //pdfFormFields.SetField("Text13", "1");
            //pdfFormFields.SetField("Text14", "1");
            //pdfFormFields.SetField("Text15", "1");
            //pdfFormFields.SetField("Text16", "1");
            //pdfFormFields.SetField("Text17", "Observación 1");
            //pdfFormFields.SetField("Text18", "1");
            //pdfFormFields.SetField("Text19", "1");
            //pdfFormFields.SetField("Text20", "1");
            //pdfFormFields.SetField("Text21", "1");
            //pdfFormFields.SetField("Text22", "Observación 2");
            //pdfFormFields.SetField("Text23", "1");
            //pdfFormFields.SetField("Text24", "1");
            //pdfFormFields.SetField("Text25", "1");
            //pdfFormFields.SetField("Text26", "1");
            //pdfFormFields.SetField("Text27", "Observación 3");
            //pdfFormFields.SetField("Text28", "1");
            //pdfFormFields.SetField("Text29", "1");
            //pdfFormFields.SetField("Text30", "1");
            //pdfFormFields.SetField("Text31", "1");
            //pdfFormFields.SetField("Text32", "Observación 4");
            //pdfFormFields.SetField("Text33", "1");
            //pdfFormFields.SetField("Text34", "1");
            //pdfFormFields.SetField("Text35", "1");
            //pdfFormFields.SetField("Text36", "1");
            //pdfFormFields.SetField("Text37", "Observación 5");
            //pdfFormFields.SetField("Text38", "1");
            //pdfFormFields.SetField("Text39", "1");
            //pdfFormFields.SetField("Text40", "1");
            //pdfFormFields.SetField("Text41", "1");
            //pdfFormFields.SetField("Text42", "Observación 6");
            //pdfFormFields.SetField("Text43", "1");
            //pdfFormFields.SetField("Text44", "1");
            //pdfFormFields.SetField("Text45", "1");
            //pdfFormFields.SetField("Text46", "1");
            //pdfFormFields.SetField("Text47", "1");
            //pdfFormFields.SetField("Text48", "1");
            //pdfFormFields.SetField("Text49", "1");
            //pdfFormFields.SetField("Text50", "1");
            //pdfFormFields.SetField("Text51", "1");
            //pdfFormFields.SetField("Text52", "1");
            //pdfFormFields.SetField("Text53", "1");
            //pdfFormFields.SetField("Text54", "1");
            //pdfFormFields.SetField("Text55", "1");
            //pdfFormFields.SetField("Text56", "1");
            //pdfFormFields.SetField("Text57", "1");
            //pdfFormFields.SetField("Text58", "1");
            //pdfFormFields.SetField("Text59", "1");
            //pdfFormFields.SetField("Text60", "1");
            //pdfFormFields.SetField("Text61", "1");
            //pdfFormFields.SetField("Text62", "1");
            //pdfFormFields.SetField("Text63", "1");
            //pdfFormFields.SetField("Text64", "1");
            //pdfFormFields.SetField("Text65", "1");
            //pdfFormFields.SetField("Text66", "1");
            //pdfFormFields.SetField("Text67", "1");
            //pdfFormFields.SetField("Text68", "1");
            //pdfFormFields.SetField("Text69", "1");
            //pdfFormFields.SetField("Text70", "1");
            //pdfFormFields.SetField("Text71", "1");
            //pdfFormFields.SetField("Text72", "1");
            //pdfFormFields.SetField("Text73", "1");
            //pdfFormFields.SetField("Text74", "1");
            //pdfFormFields.SetField("Text75", "1");
            //pdfFormFields.SetField("Text76", "1");
            //pdfFormFields.SetField("Text77", "1");
            //pdfFormFields.SetField("Text78", "1");
            //pdfFormFields.SetField("Text79", "1");
            //pdfFormFields.SetField("Text80", "1");
            //pdfFormFields.SetField("Text81", "1");
            //pdfFormFields.SetField("Text82", "1");
            //pdfFormFields.SetField("Text83", "1");
            //pdfFormFields.SetField("Text84", "1");
            //pdfFormFields.SetField("Text85", "1");
            //pdfFormFields.SetField("Text86", "1");
            //pdfFormFields.SetField("Text87", "1");
            //pdfFormFields.SetField("Text88", "1");
            //pdfFormFields.SetField("Text89", "1");
            //pdfFormFields.SetField("Text90", "1");


            // set form pdfFormFields

            // The first worksheet and W-4 form

            //pdfFormFields.SetField("f1_01(0)", "1");

            //pdfFormFields.SetField("f1_02(0)", "1");

            //pdfFormFields.SetField("f1_03(0)", "1");

            //pdfFormFields.SetField("f1_04(0)", "8");

            //pdfFormFields.SetField("f1_05(0)", "0");

            //pdfFormFields.SetField("f1_06(0)", "1");

            //pdfFormFields.SetField("f1_07(0)", "16");

            //pdfFormFields.SetField("f1_08(0)", "28");

            //pdfFormFields.SetField("f1_09(0)", "Franklin A.");

            //pdfFormFields.SetField("f1_10(0)", "Benefield");

            //pdfFormFields.SetField("f1_11(0)", "532");

            //pdfFormFields.SetField("f1_12(0)", "12");

            //pdfFormFields.SetField("f1_13(0)", "1234");



            // The form's checkboxes

            //pdfFormFields.SetField("c1_01(0)", "0");

            //pdfFormFields.SetField("c1_02(0)", "Yes");

            //pdfFormFields.SetField("c1_03(0)", "0");

            //pdfFormFields.SetField("c1_04(0)", "Yes");



            //// The rest of the form pdfFormFields

            //pdfFormFields.SetField("f1_14(0)", "100 North Cujo Street");

            //pdfFormFields.SetField("f1_15(0)", "Nome, AK  67201");

            //pdfFormFields.SetField("f1_16(0)", "9");

            //pdfFormFields.SetField("f1_17(0)", "10");

            //pdfFormFields.SetField("f1_18(0)", "11");

            //pdfFormFields.SetField("f1_19(0)", "Walmart, Nome, AK");

            //pdfFormFields.SetField("f1_20(0)", "WAL666");

            //pdfFormFields.SetField("f1_21(0)", "AB");

            //pdfFormFields.SetField("f1_22(0)", "4321");



            //// Second Worksheets pdfFormFields

            //// In order to map the fields, I just pass them a sequential

            //// number to mark them; once I know which field is which, I

            //// can pass the appropriate value

            //pdfFormFields.SetField("f2_01(0)", "1");

            //pdfFormFields.SetField("f2_02(0)", "2");

            //pdfFormFields.SetField("f2_03(0)", "3");

            //pdfFormFields.SetField("f2_04(0)", "4");

            //pdfFormFields.SetField("f2_05(0)", "5");

            //pdfFormFields.SetField("f2_06(0)", "6");

            //pdfFormFields.SetField("f2_07(0)", "7");

            //pdfFormFields.SetField("f2_08(0)", "8");

            //pdfFormFields.SetField("f2_09(0)", "9");

            //pdfFormFields.SetField("f2_10(0)", "10");

            //pdfFormFields.SetField("f2_11(0)", "11");

            //pdfFormFields.SetField("f2_12(0)", "12");

            //pdfFormFields.SetField("f2_13(0)", "13");

            //pdfFormFields.SetField("f2_14(0)", "14");

            //pdfFormFields.SetField("f2_15(0)", "15");

            //pdfFormFields.SetField("f2_16(0)", "16");

            //pdfFormFields.SetField("f2_17(0)", "17");

            //pdfFormFields.SetField("f2_18(0)", "18");

            //pdfFormFields.SetField("f2_19(0)", "19");



            //// report by reading values from completed PDF

            //string sTmp = "W-4 Completed for " + pdfFormFields.GetField("f1_09(0)") + " " + pdfFormFields.GetField("f1_10(0)");

            //MessageBox.Show(sTmp, "Finished");



            // flatten the form to remove editting options, set it to false

            // to leave the form open to subsequent manual edits

            pdfStamper.FormFlattening = false;



            // close the pdf

            pdfStamper.Close();

            Console.ReadLine();
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
