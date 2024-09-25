using ExcelProject.NET8.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using ExcelProject.NET8.Data;

namespace ExcelProject.NET8.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public ReportController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        // Export data to Excel
        [HttpGet("export-to-excel")]
        public IActionResult ExportToExcel()
        {
            var data = new List<myData>
            {
                new myData { Id = 1, Name = "John Doe", Value = 100 },
                new myData { Id = 2, Name = "Jane Smith", Value = 200 },
                new myData { Id = 3, Name = "Alice", Value = 300 }
            };

            var fileContent = _excelService.GenerateExcel(data, "Report");

            return File(fileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
        }

        // Upload an Excel file and return its contents
        [HttpPost("upload-excel")]
        public IActionResult UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded");

            var data = _excelService.ReadExcelFile<myData>(file);

            return Ok(data);
        }
    }
}
