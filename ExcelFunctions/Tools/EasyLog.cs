using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    public class EasyLog: ILog, IDisposable
    {
        System.IO.StreamWriter _loggerFile = null;
        int _errorsCount = 0;
        string _name;

        public EasyLog(string appName)
        {
            for(int i = 0; i < 20; )
            {
                try
                {
                    _name = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + appName + (i == 0? "":i.ToString()) + ".txt";
                    _loggerFile = new System.IO.StreamWriter(_name, true);
                    break;
                }
                catch(Exception ex)
                {
                    Debug.Assert(false);
                    ++i;
                }
            }
        }

        public string GetFileName()
        {
            return _name;
        }

        public int GetErrorsCount()
        {
            return _errorsCount;
        }

        public void ResetErrorsCount()
        {
            _errorsCount = 0;
        }

        public void BeginSession(string operationName)
        {
            Info($"*** Begin {operationName}: {DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss")} ***");
        }

        public void EndSession()
        {
            Info($"***End: {DateTime.Now.ToString("yyyy.MM.dd hh:mm:ss")} ***" + Environment.NewLine);
        }

        public void Info(string message)
        {
            if(_loggerFile != null)
                _loggerFile.WriteLine(message);
            Debug.WriteLine(message);
        }

        public void Note(string message)
        {
            message = "Warning: " + message;
            if (_loggerFile != null)
                _loggerFile.WriteLine(message);
            Debug.WriteLine(message);

        }
        public void Warning(string message)
        {
            message = "Warning: " + message;
            if (_loggerFile != null)
                _loggerFile.WriteLine(message);
            Debug.WriteLine(message);
        }

        public void Error(string message, Exception ex = null)
        {
            message = "Error: " + message;
            if (_loggerFile != null)
                _loggerFile.WriteLine(message);
            Debug.WriteLine(message);

            if (ex != null)
            {
                message = "Exception details: " + ex.Message;
                if (_loggerFile != null)
                    _loggerFile.WriteLine(message);
                Debug.WriteLine(message);
            }

            ++_errorsCount;
        }

        public void Flush()
        {
            if (_loggerFile != null)
                _loggerFile.Flush();
        }
        public void Dispose()
        {
            Flush();
            if(_loggerFile != null)
                _loggerFile.Dispose();
        }
    }
}
