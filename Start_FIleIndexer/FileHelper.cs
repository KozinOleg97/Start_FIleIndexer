using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using System.Configuration;

namespace Start_FIleIndexer
{
    class FileHelper
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private string pathToOldFiles;
        private string pathToNewFiles;
        private string extentionPattern = "*";

        private List<FileInfo> allFilesOld;
        private List<FileInfo> allFilesNew;
        
        //Constructor read file info from both (old & new) dirs
        public FileHelper()
        {
            //read settings from .config
            this.pathToOldFiles = ConfigurationManager.AppSettings["pathToOldFiles"];
            this.PathToNewFiles = ConfigurationManager.AppSettings["pathToNewFiles"];
            this.extentionPattern = ConfigurationManager.AppSettings["extentionPattern"];

            DirectoryInfo dirOld=null;
            DirectoryInfo dirNew=null;

            try
            {
                //read dirs
                dirOld = new DirectoryInfo(pathToOldFiles);
                dirNew = new DirectoryInfo(pathToNewFiles);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Read dirs info go wrong (Filehelper)");
                System.Environment.Exit(0);
                
            }


            //Get all files info with the specified extension
            allFilesOld = dirOld.GetFiles(extentionPattern).ToList();
            allFilesNew = dirNew.GetFiles(extentionPattern).ToList();            
        }

        //properties
        public List<FileInfo> AllFilesOld { get => allFilesOld; set => allFilesOld = value; }
        public string PathToOldFiles { get => pathToOldFiles; set => pathToOldFiles = value; }
        public string PathToNewFiles { get => pathToNewFiles; set => pathToNewFiles = value; }        

        //method that del file from list of OLD FILES if it exist in list NEW FILES. For files from REPORT
        public void DelleteByValue(string toFindValue, string toDelValue)
        {
            foreach (FileInfo indexedFile in allFilesNew)
            {
                if (indexedFile.Name == toFindValue)
                {
                    foreach(FileInfo notIndexedFile in allFilesOld)
                    {
                        if (notIndexedFile.Name == toDelValue)
                        {
                            allFilesOld.Remove(notIndexedFile);
                            return;
                        }
                    }
                }
            }


        }

    }
}
