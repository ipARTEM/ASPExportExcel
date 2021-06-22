using ASPExportExcel.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ASPExportExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult ExportExcel()
        {
            using (var workbook=new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Students");
                ws.Range("A2:E2").Merge();
                ws.Cell(1, 1).Value = "Report";
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Cell(1, 1).Style.Font.FontSize = 30;

                // Header
                ws.Cell(4, 1).Value = "DocEntry";
                ws.Cell(4, 2).Value = "Name";
                ws.Cell(4, 3).Value = "Gender";
                ws.Cell(4, 4).Value = "Phone";
                ws.Cell(4, 5).Value = "Email";
                ws.Range("A4:E4").Style.Fill.BackgroundColor = XLColor.Alizarin;

                // Connect to SQL here
                System.Data.DataTable dt = new System.Data.DataTable();
                SqlConnection con = new SqlConnection("Server=ARTEM\\ARTEM;Initial Catalog=ExpExcel;user id=sa;password=sa");
                SqlDataAdapter ad = new SqlDataAdapter("select * from Table", con);
                ad.Fill(dt);
                int i = 5;

                foreach (System.Data.DataRow row in dt.Rows)
                {
                    ws.Cell(i, 1).Value = row[0].ToString();
                    ws.Cell(i, 2).Value = row[1].ToString();
                    ws.Cell(i, 3).Value = row[2].ToString();
                    ws.Cell(i, 4).Value = row[3].ToString();
                    ws.Cell(i, 5).Value = row[4].ToString();
                    i = i + 1;
                }
                i = i - 1;
                ws.Cell("A4:E"+1).Style.Border.BottomBorder=XLBorderStyleValues.Thin;
                ws.Cell("A4:E" + 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ws.Cell("A4:E" + 1).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                ws.Cell("A4:E" + 1).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                using(var stream=new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
                        "Student.xlsx"
                        );
                }    




            }
            
        }
    }

}
