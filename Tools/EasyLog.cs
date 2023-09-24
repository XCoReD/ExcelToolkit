using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class EasyLog: ILog, IDisposable
    {
        System.IO.StreamWriter _loggerFile = null;

        public EasyLog()
        {
            string name = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\easylog.txt";
            _loggerFile = new System.IO.StreamWriter(name);
        }
        public void Info(string message)
        {
            _loggerFile.WriteLine(message);
            Console.WriteLine(message);
        }

        public void Note(string message)
        {
            message = "Warning: " + message;
            _loggerFile.WriteLine(message);
        }
        public void Warning(string message)
        {
            message = "Warning: " + message;
            _loggerFile.WriteLine(message);
        }

        public void Error(string message, Exception ex = null)
        {
            message = "Error: " + message;
            _loggerFile.WriteLine(message);

            if (ex != null)
            {
                message = "Exception details: " + ex.Message;
                _loggerFile.WriteLine(message);
            }
        }

        public void Flush()
        {
            _loggerFile.Flush();
        }
        public void Dispose()
        {
            Flush();
        }
    }
}
