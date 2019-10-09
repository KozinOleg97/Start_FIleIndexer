namespace Start_FIleIndexer
{
    class Program
    {
        static void Main(string[] args)
        {            
            ExceleHelper excele = new ExceleHelper();
            FileHelper fileHelper = new FileHelper();

            excele.CheckFromTable(fileHelper);
            excele.AddNewFiles(fileHelper);
            excele.Save();
                                   
        }
    }
}
