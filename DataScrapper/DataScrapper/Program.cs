using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using ClosedXML.Excel;
using System.Linq;

public class Program
{
    static void Main()
    {
        var url = "https://minfin.com.ua/ua/currency/";
        using var client = new HttpClient();
        Task<string> htmlTask = client.GetStringAsync(url);
        string html = htmlTask.Result;

        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var table = doc.DocumentNode.SelectSingleNode("//table[@class='sc-1x32wa2-1 dYkgjk']");
        var rows = table.SelectNodes(".//tr");

        List<string> namesList = new List<string>();
        List<string> buyPrice = new List<string>();
        List<string> sellPrice = new List<string>();
        List<string> NBUPrice = new List<string>();

        foreach (var row in rows)
        {
            var cells = row.SelectNodes("td");
            if (cells != null)
            {
                namesList.Add(cells[0].InnerText);
                buyPrice.Add(cells[1].InnerText);
                sellPrice.Add(cells[2].InnerText);
                NBUPrice.Add(cells[3].InnerText);
            }
        }

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Products");

        worksheet.Cell(1, 1).Value = "Валюта";
        worksheet.Cell(1, 2).Value = "Купівля";
        worksheet.Cell(1, 3).Value = "Продаж";
        worksheet.Cell(1, 4).Value = "Курс НБУ";

        for (int i = 0; i < namesList.Count; i++)
        {
            worksheet.Cell(i + 2, 1).Value = namesList[i];
            worksheet.Cell(i + 2, 2).Value= buyPrice[i];
            worksheet.Cell(i + 2, 3).Value = sellPrice[i];
            worksheet.Cell(i + 2, 4).Value = NBUPrice[i];
        }

        workbook.SaveAs(@"C:\Users\User\Documents\C# output\Result.xlsx");

    }
}
