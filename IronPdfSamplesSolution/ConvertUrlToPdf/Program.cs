using System;

namespace ConvertUrlToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var render = new IronPdf.HtmlToPdf();
            var doc = render.RenderUrlAsPdf("https://www.wikipedia.org/");
            doc.SaveAs($@"{AppDomain.CurrentDomain.BaseDirectory}\wiki.pdf");
        }
    }
}
