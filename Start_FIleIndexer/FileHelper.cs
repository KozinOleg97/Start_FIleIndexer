using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Configuration;
using NLog;
using System;

namespace Start_FIleIndexer
{
    class FileHelper
    {
        // ЛОГЕР
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private string pathToOldFiles;
        private string pathToNewFiles;
        private string extentionPattern = "*";

        private List<FileInfo> allFilesOld;
        private List<FileInfo> allFilesNew;

        //Constructor read file info from both (old & new) dirs
        public FileHelper()
        {
            try
            {
                //read settings from .config
                this.pathToOldFiles = ConfigurationManager.AppSettings["pathToOldFiles"];
                this.PathToNewFiles = ConfigurationManager.AppSettings["pathToNewFiles"];
                this.extentionPattern = ConfigurationManager.AppSettings["extentionPattern"];
            }
            catch (Exception ex)
            {
                logger.Error(ex, this.ToString() + " : config reading err");
                Environment.Exit(0);
            }

            try
            {
                //read dirs
                DirectoryInfo dirOld = new DirectoryInfo(pathToOldFiles);
                DirectoryInfo dirNew = new DirectoryInfo(pathToNewFiles);


                //Get all files info with the specified extension
                AllFilesOld = dirOld.GetFiles(extentionPattern).ToList();
                allFilesNew = dirNew.GetFiles(extentionPattern).ToList();

                ExcludeTempFiles();
            }
            catch (Exception ex)
            {
                logger.Error(ex, this.ToString() + " : dir paths err");
                Environment.Exit(0);
            }
        }

        //properties
        
        public string PathToOldFiles { get => pathToOldFiles; set => pathToOldFiles = value; }
        public string PathToNewFiles { get => pathToNewFiles; set => pathToNewFiles = value; }
        public List<FileInfo> AllFilesOld { get => allFilesOld; set => allFilesOld = value; }

        //method that del file from list of OLD FILES if it exist in list NEW FILES. For files from REPORT
        public void DelleteByValue(string toFindValue_newName, string toDelValue_oldName)
        {
            foreach (FileInfo indexedFile in allFilesNew)
            {
                if (indexedFile.Name == toFindValue_newName)
                {
                    foreach (FileInfo notIndexedFile in allFilesOld)
                    {
                        if (notIndexedFile.Name == toDelValue_oldName)
                        {
                            allFilesOld.Remove(notIndexedFile);
                            return;
                        }
                    }
                }
            }


        }


        public void ExcludeTempFiles()
        {
            bool allTempRemoved = false;
            while (allTempRemoved != true)
            {
                allTempRemoved = true;
                foreach (FileInfo file in allFilesOld)
                {
                    if (file.Attributes.HasFlag(FileAttributes.Hidden))
                    {
                        allTempRemoved = false;
                        allFilesOld.Remove(file);
                        break;
                    }
                }

            }
            
            
        }

    }
}
