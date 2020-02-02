using IronXL;
using IronXL.Options;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
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
    }
}
