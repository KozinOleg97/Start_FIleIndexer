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

            /*
             1) Need to specify paths
             2) Need to specify exnetion of files
             
              */

            String pathToOld = "E:\\Programming\\Start Testing Folder\\";
            String pathToNew = "E:\\Programming\\Start Testing Folder\\New Files\\";           


            ExceleHelper excele = new ExceleHelper(pathToOld);
            FileHelper fileHelper = new FileHelper(pathToOld, pathToNew);

            excele.CheckFromTable(fileHelper);
            excele.AddNewFiles(fileHelper);
            excele.Save();
                                   
        }
    }
}
