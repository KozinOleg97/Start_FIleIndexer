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
        private string path;
        private string extentionPattern = "*.doc";

        private FileInfo[] allFiles;
        //public List<FileInfo> allFiles;

        public FileHelper(string path)
        {
            this.path = path;

            //read dir
            DirectoryInfo dir = new DirectoryInfo(path);

            //Get all files info with the specified extension
            allFiles = dir.GetFiles(extentionPattern);
        }

        public void DelleteByValue(string delValue)
        {
            for (int i = 0; i < allFiles.Length; i++)
            {
                if (allFiles[i].Name == delValue)
                {
                   
                }
            }
        }

    }
}
