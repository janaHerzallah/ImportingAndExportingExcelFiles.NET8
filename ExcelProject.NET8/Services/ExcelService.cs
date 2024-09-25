namespace ExcelProject.NET8.Services
{
    using ExcelProject.NET8.Interfaces;
    using OfficeOpenXml;
    using System.Collections.Generic;
    using System.IO;

    public class ExcelService : IExcelService
    {
        public byte[] GenerateExcel<T>(List<T> data, string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(sheetName);

                // Load data into worksheet
                worksheet.Cells["A1"].LoadFromCollection(data, true);

                // Return the Excel file as byte array
                return package.GetAsByteArray();
            }
        }


        public List<T> ReadExcelFile<T>(IFormFile file) where T : new()
        {
            var dataList = new List<T>();

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Read from the first sheet
                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;

                    // Assumes the first row is a header row
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var item = new T();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var propertyInfo = typeof(T).GetProperties()[col - 1];
                            var cellValue = worksheet.Cells[row, col].Text;
                            propertyInfo.SetValue(item, Convert.ChangeType(cellValue, propertyInfo.PropertyType));
                        }
                        dataList.Add(item);
                    }
                }
            }

            return dataList;
        }
    }

}
