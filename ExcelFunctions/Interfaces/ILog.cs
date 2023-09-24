using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFunctions
{
    public interface ILog
    {
        void Info(string message);
        void Note(string message);
        void Warning(string message);
        void Error(string message, Exception ex = null);

        void BeginSession(string operationName);
        void EndSession();

        int GetErrorsCount();
        void ResetErrorsCount();

        void Flush();
        string GetFileName();
    }
}
