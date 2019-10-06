using ExcelLibrary.SpreadSheet;
using QiHe.CodeLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Start_FIleIndexer
{
    class ExceleHelper
    {
        private string path;

        private const string excelFileName = "Report";
        private const string pathToExlFile = "Report\\";
        private const string extentiot = ".xls";

        private Workbook workbook;
        private Worksheet worksheet;

        private int lastIndex = 0;

        //Load current REPORT.xlx or create new
        public ExceleHelper(string path)
        {
            this.path = path;

            try
            {
                workbook = Workbook.Load(path + pathToExlFile + excelFileName + extentiot);
                worksheet = workbook.Worksheets[0];
            }
            catch //if file not found
            {

                workbook = new Workbook();
                worksheet = new Worksheet("Report");

                // create 100 empty cells for Excel 2010>

                worksheet.Cells[0, 0] = new Cell("Index");
                worksheet.Cells[0, 1] = new Cell("New Name");
                worksheet.Cells[0, 2] = new Cell("Old Name");
                for (int i = 1; i < 101; i++)
                {
                    worksheet.Cells[i, 0] = new Cell("");
                }




            }

        }

        public void Save()
        {
            try
            {
                workbook.Worksheets[0] = worksheet;
            }
            catch
            {
                workbook.Worksheets.Add(worksheet);
            }

            workbook.Save(path + pathToExlFile + excelFileName + extentiot);

        }

        public void AddNewFiles(FileHelper fileHelper)
        {
            string newName;
            foreach (FileInfo file in fileHelper.AllFilesOld)
            {
                lastIndex++;
                newName = String.Format("{0:d4}", lastIndex) + "_" + file.Name;

                file.CopyTo(fileHelper.PathToNewFiles + newName);

                worksheet.Cells[lastIndex, 0] = new Cell(String.Format("{0:d4}", lastIndex));
                worksheet.Cells[lastIndex, 1] = new Cell(newName);
                worksheet.Cells[lastIndex, 2] = new Cell(file.Name);

                

            }
        }

        public void CheckFromTable(FileHelper fileHelper)
        {

            // traverse rows by Index
            Cell newName;
            Cell oldName;
            for (int rowIndex = worksheet.Cells.FirstRowIndex+1; rowIndex <= worksheet.Cells.LastRowIndex; rowIndex++)
            {
                Row row = worksheet.Cells.GetRow(rowIndex);

                newName = row.GetCell(1);
                oldName = row.GetCell(2);

                if (row.GetCell(0).Value.ToString() != "")
                {
                    fileHelper.DelleteByValue(newName.Value.ToString(), oldName.Value.ToString());
                    lastIndex++;
                }



                /*for (int colIndex = row.FirstColIndex;  colIndex <= row.LastColIndex; colIndex++)
                {
                    Cell cell = row.GetCell(colIndex);
                    Console.WriteLine( cell.Value.ToString());
                }*/
            }

        }


    }
}
