using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DocListManager
{
    public sealed class Logger
    {
        private static readonly Logger instance = new Logger();
        //has to be adjusted for user input, or will be stored in execution folder
        private static string logDir = null; //where to save logs ?
        private static string logFile = null; //where to save logs ?

        public enum LogLevel
        {
            Error,
            Note,
            Warning
        }

        public static string CreateLogFile(string logDir)
        {
            try
            {
                // Generate log file name with current date
                string logFileName = $"log_{DateTime.Now:dd-MM-yyyy}.txt";
                string logFilePath = Path.Combine(logDir, logFileName);

                // Create the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine("Log file created.");
                }

                return logFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Could not create log file. Details: {ex.Message}");
                return null;
            }
        }

        // Initialize method to ensure logger is constructed
        public static void Initialize(string logDir)
        {
            // This method can be expanded if initialization requires additional steps
            // Currently, it simply calls the private constructor to ensure the logger is instantiated
            _ = instance;
            // Create log file in log dir
            Logger.logDir = logDir;
            logFile = CreateLogFile(logDir);
            if (logFile != null)
            {
                Console.WriteLine($"Log file created at: {logFile}");
            }
            else
            {
                Console.WriteLine("Error: Failed to create log file.");
            }
        }

        private Logger()
        {
        }

        public static Logger Instance
        {
            get { return instance; }
        }

        public void LogWrite(string logMessage)
        {            
            try
            {

                if (File.Exists(logFile) != true)
                {
                    var file = File.CreateText(logFile);
                    file.Close();
                }

                using (StreamWriter w = File.AppendText(logFile))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception)
            {

            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                // Time & Date are logged together with Changetype and changes
                txtWriter.WriteLine($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")} : {logMessage}");
            }
            catch (Exception)
            {
            }
        }

    }

}
