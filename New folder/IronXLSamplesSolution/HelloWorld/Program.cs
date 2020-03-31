using IronXL;
using IronXL.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            decimal sum = sheet["A2:A4"].Sum();
            decimal avg = sheet["A2:A4"].Avg();
            decimal count = sheet["A2:A4"].Count();

            decimal min = sheet["A1:A4"].Min();
            decimal max = sheet["A1:A4"].Max();

            bool max2 = sheet["A1:A4"].Max(c => c.IsFormula);
            Console.WriteLine(max2);
            Console.WriteLine(min);
            Console.WriteLine(max);
            Console.WriteLine(sum);
            Console.WriteLine(avg);
            Console.WriteLine(count);
            Console.ReadLine();

        }
        static void CreateExcelFile()
        {
            var newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
            newXLFile.Metadata.Title = "IronXL New File";
            var newWorkSheet = newXLFile.CreateWorkSheet("1stWorkSheet");
            newWorkSheet["A1"].Value = "Hello World";
            newWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            newWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
        }

        static void LoadXlsx()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            var sheet = workbook.WorkSheets.First();
            var cell = sheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
        static void LoadCSV()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\CSVList.csv");
            var sheet = workbook.WorkSheets.First();
            var cell = sheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
        static void LoadXML()
        {
            var xmldataset = new DataSet();
            xmldataset.ReadXml($@"{Directory.GetCurrentDirectory()}\Files\CountryList.xml");
            var workbook = IronXL.WorkBook.Load(xmldataset);
            var sheet = workbook.WorkSheets.First();
            var cell = sheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
        static void LoadJSON()
        {
            var jsonFile = new StreamReader($@"{Directory.GetCurrentDirectory()}\Files\CountriesList.json");
            var countryList = Newtonsoft.Json.JsonConvert.DeserializeObject<CountryModel[]>(jsonFile.ReadToEnd());
            var xmldataset = countryList.ToDataSet();
            var workbook = IronXL.WorkBook.Load(xmldataset);
            var sheet = workbook.WorkSheets.First();
        }
        static void SaveXL()
        {
            var newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
            newXLFile.Metadata.Title = "IronXL New File";
            var newWorkSheet = newXLFile.CreateWorkSheet("1stWorkSheet");
            newWorkSheet["A1"].Value = "Hello World";
            newWorkSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
            newWorkSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
            newXLFile.ExportToHtml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldHTML.HTML");
            //newXLFile.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
            //newXLFile.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
            //newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldCSV.csv",delimiter:"|");
            //newXLFile.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
        }
        static void orderRange()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            sheet["A1:A6"].SortAscending(); //or use > sheet["A1:A4"].SortDescending(); to order descending
            workbook.SaveAs("SortedSheet.xlsx");

        }

        static void setCellFormula()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            int i = 1;
            foreach (var cell in sheet["B1:B6"])
            {
                cell.Formula = "=IF(A" + i + ">=20,\" Pass\" ,\" Fail\" )";
                i++;
            }
            workbook.SaveAs("testFormula.xlsx");

        }

        static void ReadRangeFormulas()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\testFormula.xlsx");
            var sheet = workbook.WorkSheets.First();
            foreach (var cell in sheet["B1:B6"])
            {
                Console.WriteLine(cell.Formula);
            }
            Console.ReadKey();

        }

        static void TrimFormula()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\TestTrim.xlsx");
            var sheet = workbook.WorkSheets.First();
            int i = 1;
            foreach (var cell in sheet["f1:f4"])
            {
                cell.Formula = "=trim(D" + i + ")";
                i++;
            }
            workbook.SaveAs("TrimFile.xlsx");

        }
        static void WorkWithSpecificSheet()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet3");
            var range = sheet["A1:D5"];
            foreach (var cell in range)
            {
                Console.WriteLine(cell.Text);
            }

        }
        static void AddNewSheetToWorkbook()
        {
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            var newSheet = workbook.CreateWorkSheet("Sheet2");
            newSheet["A1"].Value = "Hello World";
            workbook.SaveAs("NewFile.xlsx");

        }
        static void FillDbTableFromSheet()
        {

            testDBEntities dbContext = new testDBEntities();

            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet sheet = workbook.GetWorkSheet("Sheet3");

            System.Data.DataTable dataTable = sheet.ToDataTable(true);


            foreach (DataRow row in dataTable.Rows)
            {
                Country c = new Country();

                c.CountryName = row[1].ToString();
                dbContext.Countries.Add(c);

            }

            dbContext.SaveChanges();
        }
        static void FillSheetFromDb()
        {


            testDBEntities dbContext = new testDBEntities();

            WorkBook workbook = WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");

            WorkSheet sheet = workbook.CreateWorkSheet("FromDb");

            List<Country> countryList = dbContext.Countries.ToList();
            sheet.SetCellValue(0, 0, "Id");
            sheet.SetCellValue(0, 1, "Country Name");
            int row = 1;
            foreach (var item in countryList)
            {

                sheet.SetCellValue(row, 0, item.id);
                sheet.SetCellValue(row, 1, item.CountryName);
                row++;

            }
            workbook.SaveAs("FilledFile.xlsx");
        }
    }

}
