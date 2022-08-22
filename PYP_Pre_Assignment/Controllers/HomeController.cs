using Microsoft.AspNetCore.Mvc;
using System.ComponentModel;
using System;
using OfficeOpenXml;
using PYP_Pre_Assignment.Models;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using ClosedXML.Excel;

namespace PYP_Pre_Assignment.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private readonly PYPDbContext _context;
        private readonly IConfiguration _config;
        private readonly IWebHostEnvironment _env;

        public HomeController(PYPDbContext context, IConfiguration config)
        {
            _context = context;
            _config = config;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadData(IFormFile file)
        {
            string fileExt = Path.GetExtension(file.FileName);

            if (!(fileExt == ".xls" || fileExt == ".xlsx"))
                return BadRequest("only excell");

            if (file.Length / 1024 > 5120) return BadRequest("only 5 mb");

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowcount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowcount; row++)
                    {
                        XLSFile xlsfile = new();

                        xlsfile.Segment = worksheet.Cells[row, 1].Value.ToString()?.Trim();
                        xlsfile.Country = worksheet.Cells[row, 2].Value.ToString()?.Trim();
                        xlsfile.Product = worksheet.Cells[row, 3].Value.ToString()?.Trim();
                        xlsfile.DiscountBand = worksheet.Cells[row, 4].Value.ToString().Trim();
                        xlsfile.UnitsSold = double.Parse(worksheet.Cells[row, 5].Value.ToString().Trim());
                        xlsfile.ManufacturingPrice = double.Parse(worksheet.Cells[row, 6].Value.ToString().Trim());
                        xlsfile.SalePrice = double.Parse(worksheet.Cells[row, 7].Value.ToString().Trim());
                        xlsfile.GrossSales = double.Parse(worksheet.Cells[row, 8].Value.ToString().Trim());
                        xlsfile.Discounts = double.Parse(worksheet.Cells[row, 9].Value.ToString().Trim());
                        xlsfile.Sales = double.Parse(worksheet.Cells[row, 10].Value.ToString().Trim());
                        xlsfile.COGS = double.Parse(worksheet.Cells[row, 11].Value.ToString().Trim());
                        xlsfile.Profit = double.Parse(worksheet.Cells[row, 12].Value.ToString().Trim());
                        xlsfile.Date = DateTime.Parse(worksheet.Cells[row, 13].Value.ToString().Trim());

                        await _context.XLSFiles.AddAsync(xlsfile);
                    }
                }
            }

            await _context.SaveChangesAsync();
            return Ok();
        }


        [HttpGet("check")]
        public IActionResult GetData([FromQuery] DataRequest dataRequest)
        {

            var IsEmail = dataRequest.AcceptorEmail.Split("@")[1] == "code.edu.az";

            if (!IsEmail) return BadRequest("Email only code.edu.az");

            DateTime startDate = dataRequest.StartDate;

            DateTime endDate = dataRequest.EndDate;

            string email = dataRequest.AcceptorEmail;

            var query = _context.XLSFiles.Where(d => d.Date >= startDate && d.Date <= endDate);

            var datas = query.ToList();
            var mergedList = new List<DataResponse>();

            switch (dataRequest.Filter)
            {
                case FilterEnum.Segment:
                    mergedList = datas.GroupBy(x => x.Segment).Select(g => new DataResponse
                                         {
                                             FilterName = g.Key,
                                             Discount = g.Sum(x => x.Discounts),
                                             Profit = g.Sum(x => x.Profit),
                                             Sale = g.Sum(x => x.Sales),
                                             TotalCount = g.Count()
                                         })
                                         .ToList();

                    break;
                case FilterEnum.Country:
                    mergedList = datas.GroupBy(x => x.Country).Select(g => new DataResponse
                                       {
                                           FilterName = g.Key,
                                           Discount = g.Sum(x => x.Discounts),
                                           Profit = g.Sum(x => x.Profit),
                                           Sale = g.Sum(x => x.Sales),
                                           TotalCount = g.Count()
                                       })
                                         .ToList();

                    break;
                case FilterEnum.Product:
                    mergedList = datas.GroupBy(x => x.Product).Select(g => new DataResponse
                         {
                             FilterName = g.Key,
                             Discount = g.Sum(x => x.Discounts),
                             Profit = g.Sum(x => x.Profit),
                             Sale = g.Sum(x => x.Sales),
                             TotalCount = g.Count()
                         })
                                         .ToList();
                    break;
                case FilterEnum.Discount:
                    mergedList = datas.GroupBy(x => x.Product).Select(g => new DataResponse
                                       {
                                           FilterName = g.Key,
                                           Discount = g.Sum(x => x.Discounts),
                                           Profit = g.Sum(x => x.Profit),
                                           Sale = g.Sum(x => x.Sales),
                                           TotalCount = g.Count()
                                       })
                                       .ToList();
                    break;
                default:
                    break;
            }

            string excelName = $"{dataRequest.Filter.ToString()}_report.xlsx";

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Commerces");


            worksheet.Cell(2, 1).Value = $"Report date :";
            worksheet.Cell(2, 2).Value = DateTime.Now.ToString("g");
            worksheet.Cell(2, 2).DataType = XLDataType.DateTime;

            worksheet.Cell(3, 1).Value = "Filter date";
            worksheet.Cell(3, 2).Value = dataRequest.StartDate.ToString("g");
            worksheet.Cell(3, 2).DataType = XLDataType.DateTime;

            worksheet.Cell(3, 3).Value = "-";
            worksheet.Cell(3, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(3, 4).Value = dataRequest.EndDate.ToString("g");
            worksheet.Cell(3, 4).DataType = XLDataType.DateTime;

            var currentRow = 7;

            worksheet.Row(currentRow).Height = 25.0;
            worksheet.Row(currentRow).Style.Font.Bold = true;
            worksheet.Row(currentRow).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Cell(currentRow, 1).Value = "FilterName";
            worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 2).Value = "Discount";
            worksheet.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 3).Value = "Profit";
            worksheet.Cell(currentRow, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 4).Value = "Sale";
            worksheet.Cell(currentRow, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell(currentRow, 5).Value = "TotalCount";
            worksheet.Cell(currentRow, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            foreach (var item in mergedList)
            {
                currentRow++;

                worksheet.Cell(currentRow, 1).Value = item.FilterName;
                worksheet.Cell(currentRow, 2).Value = item.Discount;
                worksheet.Cell(currentRow, 3).Value = item.Profit;
                worksheet.Cell(currentRow, 4).Value = item.Sale;
                worksheet.Cell(currentRow, 5).Value = item.TotalCount;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            EmailSender emailService = new EmailSender(_config.GetSection("ConfirmationParams:Email").Value, _config.GetSection("ConfirmationParams:Password").Value);
            emailService.SendEmail(email, "excell", "bax", excelName, content);

            Ok("Gonderildi");
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }
    }
}
