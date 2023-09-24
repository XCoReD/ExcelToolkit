using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tools
{
    public interface ILog
    {
        void Info(string message);
        void Note(string message);
        void Warning(string message);
        void Error(string message, Exception ex = null);
    }
}
