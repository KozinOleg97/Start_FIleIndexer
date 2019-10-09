using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Configuration;

namespace Start_FIleIndexer
{
    class FileHelper
    {
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

            //read dirs
            DirectoryInfo dirOld = new DirectoryInfo(pathToOldFiles);
            DirectoryInfo dirNew = new DirectoryInfo(pathToNewFiles);


            //Get all files info with the specified extension
            AllFilesOld1  = dirOld.GetFiles(extentionPattern).ToList();
            allFilesNew = dirNew.GetFiles(extentionPattern).ToList();            
        }

        //properties
        public List<FileInfo> AllFilesOld { get => AllFilesOld1; set => AllFilesOld1 = value; }
        public string PathToOldFiles { get => pathToOldFiles; set => pathToOldFiles = value; }
        public string PathToNewFiles { get => pathToNewFiles; set => pathToNewFiles = value; }
        public List<FileInfo> AllFilesOld1 { get => allFilesOld; set => allFilesOld = value; }

        //method that del file from list of OLD FILES if it exist in list NEW FILES. For files from REPORT
        public void DelleteByValue(string toFindValue, string toDelValue)
        {
            foreach (FileInfo indexedFile in allFilesNew)
            {
                if (indexedFile.Name == toFindValue)
                {
                    foreach(FileInfo notIndexedFile in AllFilesOld1)
                    {
                        if (notIndexedFile.Name == toDelValue)
                        {
                            AllFilesOld1.Remove(notIndexedFile);
                            return;
                        }
                    }
                }
            }


        }

    }
}
