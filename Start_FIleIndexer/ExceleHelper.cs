using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Start_FIleIndexer
{
    class ExceleHelper
    {
        private string path;

        private const string ExcelFileName = "Report";
        private const string pathToExlFile = "Report\\";


        public ExceleHelper(string path)
        {
            this.path = path;

            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("Report");

            // create 100 empty cells for Excel 2010>
            for (int i = 0; i < 100; i++)
                worksheet.Cells[i, 0] = new Cell("");

            workbook.Worksheets.Add(worksheet);
            workbook.Save(path + pathToExlFile + ExcelFileName + ".xls");
        }

        public Workbook LoadFromFile()
        {

            return Workbook.Load("path");
        }




    }
}
