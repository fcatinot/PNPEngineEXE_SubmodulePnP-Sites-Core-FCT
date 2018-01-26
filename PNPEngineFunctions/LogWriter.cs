using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace PNPEngineFunctions
{
    public class LogWriter
    {
        private static volatile LogWriter _instance;
        private static readonly object Lock = new object();

        private LogWriter()
        {
        }

        public static LogWriter Current
        {
            get
            {
                if (_instance == null)
                {
                    lock (Lock)
                    {
                        if (_instance == null)
                            _instance = new LogWriter();
                    }
                }

                return _instance;
            }
        }

        private StringBuilder fileOutput = new StringBuilder();

        public void WriteLine(string toWrite)
        {
            Console.WriteLine(toWrite);
            fileOutput.AppendLine(toWrite);
        }

        public void UpdateLogFile(string logsFileName)
        {
            using (
                StreamWriter writer =
                    new StreamWriter(
                        Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" +
                        logsFileName, false))
            {
                writer.Write(fileOutput.ToString());
            }
        }
    }
}
