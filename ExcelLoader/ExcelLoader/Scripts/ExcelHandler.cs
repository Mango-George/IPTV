using System;
using System.IO;

namespace ExcelLoader.Scripts
{
    public class ExcelHandler
    {
        public void Export()
        {
            DirectoryInfo directoryInfo = new DirectoryInfo("Datas");
            FileInfo[] fileInfos = directoryInfo.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
            Console.Write(fileInfos.Length);
            foreach (FileInfo fileInfo in fileInfos)
            {
                Console.WriteLine(fileInfo.FullName);
            }
        }
    }
}