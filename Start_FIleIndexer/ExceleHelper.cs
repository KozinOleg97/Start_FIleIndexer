using IronXL;
using IronXL.Options;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Start_FIleIndexer
{
    class ExceleHelper
    {
        private string excelFileName;
        private string pathToExlFile;
        private string extention = ".xls";
        private int mainFontHight;

        private WorkBook workbook;
        private WorkSheet worksheet;

        //Index of last not empty raw
        private int lastIndex = 0;

        //Load current REPORT.xls or create new
        public ExceleHelper()
        {
            //read settings from .config
            this.pathToExlFile = ConfigurationManager.AppSettings["pathToExlFile"];
            this.excelFileName = ConfigurationManager.AppSettings["excelFileName"];
            this.mainFontHight = Int32.Parse(ConfigurationManager.AppSettings["mainFontHight"]);

            try
            {
                workbook = WorkBook.Load(pathToExlFile + excelFileName + extention);
                worksheet = workbook.WorkSheets.First();
            }
            catch //if file not found
            {
                workbook = WorkBook.Load("Sample.xls");
                worksheet = workbook.WorkSheets.First();
            }

        }

        //Add result worksheet to current workbook & save it
        public void Save()
        {

            worksheet.Columns[0].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[1].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[2].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[2].Style.RightBorder.Type = IronXL.Styles.BorderType.Thin;

            worksheet.Columns[0].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[0].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;

            worksheet.Columns[1].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[1].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;

            worksheet.Columns[2].Style.TopBorder.Type = IronXL.Styles.BorderType.Thin;
            worksheet.Columns[2].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thin;

            worksheet.Style.Font.Height = (short)mainFontHight;


            worksheet.Rows[0].Style.Font.Height = 10;

            try
            {                
                workbook.SaveAs(pathToExlFile + excelFileName + extention);
            }
            catch
            {
                //if no worksheet
                //workbook.Worksheets.Add(worksheet);
            }

            //workbook.Save(pathToExlFile + excelFileName + extention);

        }

        //Add files to REPORT & create NEW FILE with INDEX
        public void AddNewFiles(FileHelper fileHelper)
        {
            string newName;
            foreach (FileInfo file in fileHelper.AllFilesOld)
            {

                newName = String.Format("{0:d5}", lastIndex) + "_" + file.Name;

                file.CopyTo(fileHelper.PathToNewFiles + newName);

                worksheet["A" + (lastIndex + 1)].Value = String.Format("{0:d5}", lastIndex);
                worksheet["B" + (lastIndex + 1)].Value = file.Name;
                worksheet["c" + (lastIndex + 1)].Value = newName;

                worksheet["B" + (lastIndex + 1)].Style.WrapText = true;
                worksheet["c" + (lastIndex + 1)].Style.WrapText = true;

                lastIndex++;
            }
        }

        //iterates through REPORT & remove existing files from consideration
        public void CheckFromTable(FileHelper fileHelper)
        {
            Range newName;
            Range oldName;

            lastIndex = worksheet.Rows.Count();

            bool firstRowPass = true;
            foreach (var row in worksheet.Rows)
            {
                if (firstRowPass) { firstRowPass = false; continue; }

                newName = row.Columns[2];
                oldName = row.Columns[1];

                fileHelper.DelleteByValue(newName.StringValue, oldName.StringValue);
            }          

        }


    }
}
