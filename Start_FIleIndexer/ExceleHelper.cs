using ExcelLibrary.SpreadSheet;
using QiHe.CodeLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Start_FIleIndexer
{
    class ExceleHelper
    {      
        private string excelFileName;
        private string pathToExlFile;
        private string extention = ".xls";

        private Workbook workbook;
        private Worksheet worksheet;

        //Index of last not empty raw
        private int lastIndex = 0;

        //Load current REPORT.xls or create new
        public ExceleHelper()
        {
            //read settings from .config
            this.pathToExlFile = ConfigurationManager.AppSettings["pathToExlFile"];
            this.excelFileName = ConfigurationManager.AppSettings["excelFileName"];
            

            try
            {
                workbook = Workbook.Load(pathToExlFile + excelFileName + extention);
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

        //Add result worksheet to current workbook & save it
        public void Save()
        {
            try
            {
                workbook.Worksheets[0] = worksheet;
            }
            catch
            {
                //if no worksheet
                workbook.Worksheets.Add(worksheet);
            }

            workbook.Save(pathToExlFile + excelFileName + extention);

        }

        //Add files to REPORT & create NEW FILE with INDEX
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

        //iterates through REPORT & remove existing files from consideration
        public void CheckFromTable(FileHelper fileHelper)
        {            
            Cell newName;
            Cell oldName;
            for (int rowIndex = worksheet.Cells.FirstRowIndex + 1; rowIndex <= worksheet.Cells.LastRowIndex; rowIndex++)
            {
                Row row = worksheet.Cells.GetRow(rowIndex);

                newName = row.GetCell(1);
                oldName = row.GetCell(2);

                if (row.GetCell(0).Value.ToString() != "")
                {
                    fileHelper.DelleteByValue(newName.Value.ToString(), oldName.Value.ToString());
                    lastIndex++;
                }

            }

        }


    }
}
