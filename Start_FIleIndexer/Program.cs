using NLog;

namespace Start_FIleIndexer
{
    class Program
    {
        // ЛОГЕР
        private static Logger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            logger.Info("Programm started!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

            ExceleHelper excele = new ExceleHelper();
            FileHelper fileHelper = new FileHelper();

            excele.CheckFromTable(fileHelper);
            excele.AddNewFiles(fileHelper);
            excele.Save();

            logger.Info("Program has terminated successfully!!!!!!!!!!!!!!!!!!!!!!");
        }
    }
}
