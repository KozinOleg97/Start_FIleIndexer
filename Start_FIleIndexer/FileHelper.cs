using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Start_FIleIndexer
{
    class FileHelper
    {
        private string pathToOldFiles;
        private string pathToNewFiles;
        private string extentionPattern = "*.doc";

        private List<FileInfo> allFilesOld;
        private List<FileInfo> allFilesNew;
        

        public FileHelper(string pathToOldFiles, string pathToNewFiles)
        {
            this.PathToOldFiles = pathToOldFiles;
            this.PathToNewFiles = pathToNewFiles;

            //read dir
            DirectoryInfo dirOld = new DirectoryInfo(pathToOldFiles);
            DirectoryInfo dirNew = new DirectoryInfo(pathToNewFiles);


            //Get all files info with the specified extension
            allFilesOld  = dirOld.GetFiles(extentionPattern).ToList();
            allFilesNew = dirNew.GetFiles(extentionPattern).ToList();            
        }

        public List<FileInfo> AllFilesOld { get => allFilesOld; set => allFilesOld = value; }
        public string PathToOldFiles { get => pathToOldFiles; set => pathToOldFiles = value; }
        public string PathToNewFiles { get => pathToNewFiles; set => pathToNewFiles = value; }

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










         /*   for (int i = 0; i < allFilesNew.Length; i++)
            {                
                if (allFilesNew[i].Name == toFindValue)
                {
                   for (int j = 0; j < allFilesOld.Length; j++)
                    {
                        if (allFilesOld[i].Name == toDelValue)
                        {
                            allFilesOld[i] = new Fi;
                            return;
                        }
                    }
                }
            }*/
        }

    }
}
