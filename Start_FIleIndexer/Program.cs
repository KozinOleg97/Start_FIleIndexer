using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Start_FIleIndexer
{
    class Program
    {
        static void Main(string[] args)
        {

            String path = "E:\\Programming\\Start Testing Folder\\";

            //Console.WriteLine("Path:");
            //path = Console.ReadLine();            

            //creating fileName
            String newExcelFileName = "Report";// + DateTime.Now.ToString("yyyy-MM-dd (h mm ss)tt");

            ExceleHelper excele = new ExceleHelper(path);


            Workbook workbook = excele.LoadFromFile();

            //initialization & new doc
            DirectoryInfo dir = new DirectoryInfo(path);


            //write to cells
            FileInfo[] allFiles = dir.GetFiles("*.doc");

            for (int cnt = 0; cnt < allFiles.Length; cnt++)
            {

                worksheet.Cells[cnt, 0] = new Cell(allFiles[cnt].Name);
                worksheet.Cells[cnt, 1] = new Cell(cnt);
                File.Move(allFiles[cnt].FullName, path + cnt + "_" +
                    Path.GetFileNameWithoutExtension(allFiles[cnt].Name) + ".doc");

            }


            // System.IO.File.Move("oldfilename", "newfilename");

            //save
            workbook.Worksheets.Add(worksheet);
            workbook.Save(path + newExcelFileName + ".xls");



            Console.WriteLine("Done");
            Console.ReadKey();
        }
    }
}
