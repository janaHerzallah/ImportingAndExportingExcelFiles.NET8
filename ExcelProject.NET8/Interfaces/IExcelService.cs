namespace ExcelProject.NET8.Interfaces
{
    public interface IExcelService
    {
        byte[] GenerateExcel<T>(List<T> data, string sheetName);
        List<T> ReadExcelFile<T>(IFormFile file) where T : new();

    }

}
