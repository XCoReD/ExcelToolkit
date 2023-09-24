using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelFunctions
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public void OnFillByTemplateClick(IRibbonControl control)
        {
            using (EasyLog log = new EasyLog("ExcelToolkit"))
            {
                var excel = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                TemplateProcessor.Process(excel, log);
                log.Flush();
                if (log.GetErrorsCount() != 0)
                {
                    string msg = @"
One or more errors are found during processing template.
Open the log file and check details.

If you want the ExcelToolkit to open the log file now, press OK
";
                    if (MessageBox.Show(msg, "Excel Toolkit", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) == DialogResult.OK)
                    {
                        Process.Start(log.GetFileName());
                    }
                }
            }
        }

        public void OnSyncWithSharepointListClick(IRibbonControl control)
        {
            using (EasyLog log = new EasyLog("ExcelToolkit"))
            {
                var excel = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                SharepointListProcessor.Process(excel, log);
                log.Flush();
                if (log.GetErrorsCount() != 0)
                {
                    string msg = @"
One or more errors are found during processing the spreadsheet.
Open the log file and check details.

If you want the ExcelToolkit to open the log file now, press OK
";
                    if (MessageBox.Show(msg, "Excel Toolkit", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) == DialogResult.OK)
                    {
                        Process.Start(log.GetFileName());
                    }
                }
            }
        }

        public void OnUpdateButtonClick(IRibbonControl control)
        {
            var excel = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

            excel.CalculateFull();
        }

        public void OnHowToUseFunctionsClick(IRibbonControl control)
        {
            // Show Message Boxes with Excel-DNA
            // https://andysprague.com/2017/07/03/show-message-boxes-with-excel-dna/
            string msg = @"
Available functions are:

1) Obtain the currency exchange rate defined by European Central Bank \nin terms of a base currency rate on a given date
=DT.ExchangeRate(""2019/11/29"",""USD"",""EUR"")

2) Get USD rate on a given date(e.g. 2019 / 11 / 29) from NBRB web service
   = DT.ExchangeUSDRateNBRB(""2019/11/29"")

3) Get suma in cursive for a given amount in a given currency and given verb(Nominative(empty), Genitive(""RP"") or Dative(""DP""))
= DT.SumProp(12.34, ""USD"", ""DP"")
= DT.SumProp(12345.67, ""BYN"", """")

(you can remember them or use Formulas->Insert Function, then navigate to DT group)

";
            MessageBox.Show(msg, "Excel Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnAboutButtonClick(IRibbonControl control)
        {
            new AboutBox1().ShowDialog();
        }

    }
}