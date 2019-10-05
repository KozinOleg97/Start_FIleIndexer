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

        public Workbook workbook;
        public Worksheet worksheet;

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
                for (int i = 0; i < 101; i++)
                    worksheet.Cells[i, 0] = new Cell("");

                //workbook.Worksheets.Add(worksheet);
                //workbook.Save(path + pathToExlFile + excelFileName + extentiot);
            }




        }


        public void CheckFromTable(FileHelper fileHelper)
        {

            // traverse rows by Index
            Cell cell_0;
            Cell cell_1;
            for (int rowIndex = worksheet.Cells.FirstRowIndex; rowIndex <= worksheet.Cells.LastRowIndex; rowIndex++)
            {
                Row row = worksheet.Cells.GetRow(rowIndex);

                cell_0 = row.GetCell(0);
                cell_1 = row.GetCell(1);

                fileHelper.DelleteByValue(cell_1.Value.ToString());

                /*for (int colIndex = row.FirstColIndex;  colIndex <= row.LastColIndex; colIndex++)
                {
                    Cell cell = row.GetCell(colIndex);
                    Console.WriteLine( cell.Value.ToString());
                }*/
            }

        }


    }
}
