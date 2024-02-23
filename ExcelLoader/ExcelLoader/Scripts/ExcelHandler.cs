using OfficeOpenXml;
using System;
using System.IO;

namespace ExcelLoader.Scripts
{
    public class ExcelHandler
    {
        public void Export()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DirectoryInfo directoryInfo = new DirectoryInfo("Datas");
            FileInfo[] fileInfos = directoryInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
            Console.Write(fileInfos.Length);
            // foreach (FileInfo fileInfo in fileInfos)
            FileInfo fileInfo = fileInfos[0];
            {
                Console.WriteLine(fileInfo.FullName);
                using (FileStream stream = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.ReadWrite))
                {
                    try
                    {
                        ExcelPackage excelPackage = new ExcelPackage(stream);
                        ExcelWorkbook workbook = excelPackage.Workbook;
                        int sheetCount = workbook.Worksheets.Count;
                        for (int i = 0; i < sheetCount; i++)
                        {
                            ExcelWorksheet worksheet = workbook.Worksheets[i];
                            int row = worksheet.Dimension.End.Row;
                            int column = worksheet.Dimension.End.Column;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        throw;
                    }
                }
            }
        }
    }
}