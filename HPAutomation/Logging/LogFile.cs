namespace HPAutomation.Logging
{
    using System;
    using System.IO;

    public class LogFile
    {
        private static string FilePath { get; set; }

        public static void SetFilePath()
        {
            var fileName = "LogFile" + DateTime.Now.ToString("yyyyMMddTHHmmss") + ".txt";
            var directory = Directory.GetCurrentDirectory();

            var logDirectory = Path.Combine(directory, "LogFiles");

            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            FilePath = Path.Combine(logDirectory, fileName);
        }

        public static void WriteLog(string message)
        {
            Console.WriteLine(message);

            var fileStream = new FileStream(FilePath, FileMode.Append, FileAccess.Write);
            var streamWriter = new StreamWriter(fileStream);

            try
            {
                streamWriter.WriteLine(message);
            }
            finally
            {
                streamWriter.Close();
                fileStream.Close();
            }
        }
    }
}
