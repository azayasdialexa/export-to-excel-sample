using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportToExcelSample.Models;
using ExportToExcelSample.Utility;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExportToExcelSample.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View(DataStorage.GetAllAvailable());
        }
        public IActionResult Csv()
        {
            var availableMarkets = DataStorage.GetAllMarkets();
            var availableSummaries = DataStorage.GetAllAvailable();
            var builder = new StringBuilder();

            builder.AppendLine("Market Id,Address,City,State,Total Sf,Clear Ht,Doors,Avail Date,List Rep,Rep Email");

            foreach (var market in availableMarkets)
            {
                foreach (var summary in availableSummaries.Where(s => s.MarketId == market.MarketId))
                {
                    builder.AppendLine($"{summary.MarketId}," +
                        $"{summary.Address}," +
                        $"{summary.City}," +
                        $"{summary.State}," +
                        $"{summary.TotalSf}," +
                        $"{summary.ClearHt}," +
                        $"{summary.Doors}," +
                        $"{summary.AvailableDate.ToShortDateString()}," +
                        $"{summary.ListingRep}," +
                        $"{summary.RepEmail}");
                }
            }
            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "markets.csv");
        }
        public IActionResult Excel()
        {
            var availableMarkets = DataStorage.GetAllMarkets();
            var availableSummaries = DataStorage.GetAllAvailable();
            using var workbook = new XLWorkbook();

            foreach (var market in availableMarkets)
            {
                var row = 1;
                var worksheet = workbook.Worksheets.Add(market.Name).SetTabColor(XLColor.FromName(market.TabColor));

                worksheet.Row(row).Height = 25.0;
                worksheet.Row(row).Style.Font.Bold = true;
                worksheet.Row(row).Style.Font.FontColor = XLColor.Black;
                worksheet.Row(row).Style.Fill.BackgroundColor = XLColor.LightGray;
                worksheet.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;

                var col = 0;
                col++; worksheet.Cell(row, col).Value = "Market Id"; 
                col++; worksheet.Cell(row, col).Value = "Address";
                col++; worksheet.Cell(row, col).Value = "City";
                col++; worksheet.Cell(row, col).Value = "State";
                col++; worksheet.Cell(row, col).Value = "Total Sf"; 
                col++; worksheet.Cell(row, col).Value = "Clear Ht"; 
                col++; worksheet.Cell(row, col).Value = "Doors"; 
                col++; worksheet.Cell(row, col).Value = "Avail Date"; 
                col++; worksheet.Cell(row, col).Value = "List Rep";

                worksheet.RangeUsed().SetAutoFilter();

                foreach (var summary in availableSummaries.Where(s => s.MarketId == market.MarketId))
                {
                    row++;
                    worksheet.Row(row).Height = 20.0;
                    worksheet.Row(row).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    col = 0;
                    col++; worksheet.Cell(row, col).Value = summary.MarketId;
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, col).DataType = XLDataType.Number;

                    col++; worksheet.Cell(row, col).Value = summary.Address;
                    col++; worksheet.Cell(row, col).Value = summary.City;
                    col++; worksheet.Cell(row, col).Value = summary.State;
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    col++; worksheet.Cell(row, col).Value = summary.TotalSf;
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, col).DataType = XLDataType.Number;
                    worksheet.Cell(row, col).Style.NumberFormat.Format = "#,##0";
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrGreaterThan(750000).Font.FontColor = XLColor.Purple;
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrGreaterThan(750000).Font.Bold = true;
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrLessThan(250000).Font.FontColor = XLColor.Red;
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrLessThan(250000).Font.Bold = true;

                    col++; worksheet.Cell(row, col).Value = summary.ClearHt;
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, col).DataType = XLDataType.Number;

                    col++; worksheet.Cell(row, col).Value = summary.Doors;
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, col).DataType = XLDataType.Number;

                    worksheet.Column(col).AddConditionalFormat().WhenEquals(0).Font.FontColor = XLColor.Crimson;
                    worksheet.Column(col).AddConditionalFormat().WhenEquals(0).Font.Bold = true;
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrGreaterThan(100).Font.FontColor = XLColor.Purple;
                    worksheet.Column(col).AddConditionalFormat().WhenEqualOrGreaterThan(100).Font.Bold = true;

                    col++; worksheet.Cell(row, col).Value = summary.AvailableDate.ToShortDateString();
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    worksheet.Cell(row, col).DataType = XLDataType.DateTime;
                    worksheet.Column(col).AddConditionalFormat().WhenLessThan(DateTime.Now.ToOADate()).Font.FontColor = XLColor.Purple;
                    worksheet.Column(col).AddConditionalFormat().WhenLessThan(DateTime.Now.ToOADate()).Font.Bold = true;

                    col++; worksheet.Cell(row, col).Value = summary.ListingRep;
                    worksheet.Cell(row, col).Hyperlink.ExternalAddress = new Uri($"mailto:{summary.RepEmail}");
                }
                worksheet.Columns().AdjustToContents(); // NOTE: Adjust contents after adding all columns
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "markets.xlsx");
        }
    }
}
