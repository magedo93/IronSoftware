using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelloWorld
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = IronXL.WorkBook.Load(@"D:\EbtkarSolutions\IronSoftware\Tutorials\IronSoftware\IronXLSamplesSolution\Documents\Files\CSVList.csv");
            var sheet = workbook.WorkSheets.First();
            var cell = sheet["A1"].StringValue;
            Console.WriteLine(cell);
        }

        static void LoadXlsx()
        {
            var workbook = IronXL.WorkBook.Load(@"D:\EbtkarSolutions\IronSoftware\Tutorials\IronSoftware\IronXLSamplesSolution\Documents\Files\HelloWorld.xlsx");
            var sheet = workbook.WorkSheets.First();
            var cell = sheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}
