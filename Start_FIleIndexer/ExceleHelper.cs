
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.POIFS.FileSystem;
using NPOI.HPSF;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Data;
using NPOI.XSSF.UserModel;

namespace Start_FIleIndexer
{
    class ExceleHelper
    {
        private string excelFileName;
        private string pathToExlFile;
        private string extention = ".xls";
        private int mainFontHight;



        private IWorkbook workbook;
        private ISheet worksheet;
        

        //Index of last not empty raw
        private int lastIndex = 0;

        //Load current REPORT.xls or create new
        public ExceleHelper()
        {
            //read settings from .config
            this.pathToExlFile = ConfigurationManager.AppSettings["pathToExlFile"];
            this.excelFileName = ConfigurationManager.AppSettings["excelFileName"];
            this.mainFontHight = Int32.Parse(ConfigurationManager.AppSettings["mainFontHight"]);

           
            FileInfo file = new FileInfo(pathToExlFile + excelFileName + extention);
            if (file.Exists)
            {
                using (FileStream xlsFile = new FileStream(pathToExlFile + excelFileName + extention, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(xlsFile);
                }
            }
            else
            {
                using (FileStream xlsFile = new FileStream("Sample.xls", FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(xlsFile);
                }
            }

            worksheet = workbook.GetSheet("Отчёт");

        }

        //Add result worksheet to current workbook & save it
        public void Save()
        {


            using (FileStream stream = new FileStream(pathToExlFile + excelFileName + extention, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
            
        }

        //Add files to REPORT & create NEW FILE with INDEX
        public void AddNewFiles(FileHelper fileHelper)
        {

            IRow curRow = worksheet.GetRow(0);
            string newName;
            //ICell oldName;                

            foreach (FileInfo file in fileHelper.AllFilesOld)
            {
                newName = String.Format("{0:d5}", lastIndex) + "_" + file.Name;

                file.CopyTo(fileHelper.PathToNewFiles + newName);


                curRow = worksheet.CreateRow(lastIndex);
                curRow.CreateCell(0).SetCellValue(String.Format("{0:d5}", lastIndex));
                curRow.CreateCell(1).SetCellValue(file.Name);
                curRow.CreateCell(2).SetCellValue(newName);

                editCellStyle(curRow.GetCell(0));
                editCellStyle(curRow.GetCell(1));
                editCellStyle(curRow.GetCell(2));

                
                lastIndex++;
            }
        }

        //iterates through REPORT & remove existing files from consideration
        public void CheckFromTable(FileHelper fileHelper)
        {

            IRow curRow = worksheet.GetRow(0);

            ICell index = curRow.GetCell(0);
            ICell newName = curRow.GetCell(2);
            ICell oldName = curRow.GetCell(1);

            editCellStyle(index);
            editCellStyle(oldName);
            editCellStyle(newName);


            lastIndex = worksheet.LastRowNum + 1;

            for (int rowIndex = 1; rowIndex < worksheet.LastRowNum + 1; rowIndex++)
            {
                curRow = worksheet.GetRow(rowIndex);
                newName = curRow.GetCell(2);
                oldName = curRow.GetCell(1);

                fileHelper.DelleteByValue(newName.StringCellValue, oldName.StringCellValue);
            }

        }


        private void editCellStyle(ICell cell)
        {
            cell.CellStyle.WrapText = true;
            
            cell.CellStyle.BorderBottom = BorderStyle.Thin;            
            cell.CellStyle.BorderTop = BorderStyle.Thin;
            cell.CellStyle.BorderLeft = BorderStyle.Thin;
            cell.CellStyle.BorderRight = BorderStyle.Thin;
        }

    }


}