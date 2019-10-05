using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
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
            

            ExceleHelper excele = new ExceleHelper(path);
            FileHelper fileHelper = new FileHelper(path);

            excele.CheckFromTable(fileHelper);

            //Workbook workbook = excele.LoadFromFile();

            

            /*for (int cnt = 0; cnt < allFiles.Length; cnt++)
            {

                worksheet.Cells[cnt, 0] = new Cell(allFiles[cnt].Name);
                worksheet.Cells[cnt, 1] = new Cell(cnt);
                File.Move(allFiles[cnt].FullName, path + cnt + "_" +
                    Path.GetFileNameWithoutExtension(allFiles[cnt].Name) + ".doc");

            }


            // System.IO.File.Move("oldfilename", "newfilename");

            //save
            workbook.Worksheets.Add(worksheet);
            workbook.Save(path + newExcelFileName + ".xls");*/



            Console.WriteLine("Done");
            Console.ReadKey();
        }
    }
}
