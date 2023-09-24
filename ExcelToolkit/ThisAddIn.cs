using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelToolkit
{
    public partial class ThisAddIn
    {
        Settings _settings;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _settings = new Settings(Application);
            _settings.RegisterFunctions();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _settings.UnRegisterFunctions();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
