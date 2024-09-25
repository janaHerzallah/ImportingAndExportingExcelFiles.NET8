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

                // loadFromCollection method loads the list data into the worksheet starting at cell A1
                //true argument indicates that the method should use the property names of the objects in the list as the headers of the columns.

                worksheet.Cells["A1"].LoadFromCollection(data, true);

                //the byte array is used to return the Excel file to the client as a file download.
                return package.GetAsByteArray();

                /*This method takes a list of data, generates an Excel file with the data,
                 * and returns the Excel file as a byte array, which can be sent to the client as a downloadable file.*/
            }
        }

        // read the content of a file and then convert it to a list of objects of type T
        // type T means that the method can read any type of object from an Excel file
        // so we can specify the the type of object we want to read from the Excel file

        public List<T> ReadExcelFile<T>(IFormFile file) where T : new() // new() constraint means that the type T must have a parameterless constructor
            //IFormFile interface represents a file sent with an HTTP request.

        {
            var dataList = new List<T>();

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream); //Copying the Uploaded File to a MemoryStream:

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
