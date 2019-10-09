namespace Start_FIleIndexer
{
    class Program
    {
        static void Main(string[] args)
        {

            /*
             1) Need to specify paths
             2) Need to specify extensions of files
             
              */

            //String pathToOld = "E:\\Programming\\Start Testing Folder\\";
           // String pathToNew = "E:\\Programming\\Start Testing Folder\\New Files\\";           


            ExceleHelper excele = new ExceleHelper();
            FileHelper fileHelper = new FileHelper();

            excele.CheckFromTable(fileHelper);
            excele.AddNewFiles(fileHelper);
            excele.Save();
                                   
        }
    }
}
