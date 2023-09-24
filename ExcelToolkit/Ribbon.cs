using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace ExcelToolkit
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //https://coderwall.com/p/app3ya/read-excel-file-in-c
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<Param> arguments = new List<Param>(colCount);

            for (int i = 1; i <= rowCount; i++)
            {
                List<Param> values = i == 1 ? arguments : arguments.GetRange(0, arguments.Count);
                bool processRow = i > 1;
                int indexOutputDoc = 0;
                int indexOutputDocGenDate = 0;
                string templatePath = null;


                for (int j = 1; j <= colCount; j++)
                {
                    if(i == 1)
                    {
                        string name = null;
                        Range range = xlRange.Cells[i, j];
                        if (range != null && range.Value2 != null)
                        {
                            name = range.Value2.ToString();
                        }
                        Param param = new Param { name = name };
                        values.Add(param);
                    }
                    else
                    {
                        string value = null;
                        Range range = xlRange.Cells[i, j];
                        if (range != null && range.Value2 != null)
                        {
                            value = range.Value2.ToString();
                        }

                        DisplayType type = DisplayType.Text;
                        double test;
                        if(double.TryParse(value, out test))
                        {
                            string sfmt = range.NumberFormat.ToString();
                            if(sfmt != "General")
                            {
                                try
                                {
                                    DateTime dt = DateTime.FromOADate(test);
                                    if (sfmt.IndexOf("yy-mm-dd") >= 0 || sfmt.IndexOf("MM-yy") >= 0 ||
                                        sfmt.IndexOf("mmm-yy") >= 0 || sfmt.IndexOf("dd-mmm") >= 0 ||
                                        sfmt.IndexOf("dd/mm/yyyy") >= 0 || sfmt.IndexOf("yyyy mm dd") >= 0 ||
                                        sfmt.IndexOf("dd mmm, yyyy") >= 0 || sfmt.IndexOf("yyyy") >= 0)
                                    {
                                        type = DisplayType.Date;
                                    }
                                    else if (sfmt.IndexOf('0') == 0)
                                    {
                                        //number format string
                                        type = DisplayType.Number;
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"Excel value {value}, format {sfmt} - suppose a text");
                                        type = DisplayType.Text;
                                    }
                                }
                                catch (ArgumentException ex)
                                {
                                    //not date and not has number format - suppose text as well
                                    type = DisplayType.Text;
                                }
                            }
                        }

                        if (values.ElementAt(j-1).name != null)
                        {
                            //process predefined parameters
                            switch(values.ElementAt(j - 1).name.ToLower())
                            {
                                case "docgen":
                                    if(!string.IsNullOrEmpty(value))
                                    {
                                        Debug.WriteLine($"{value} is specified, skipping");
                                        processRow = false;
                                    }
                                    indexOutputDoc = j;
                                    break;
                                case "doctemplate":
                                    if (string.IsNullOrEmpty(value))
                                    {
                                        Debug.WriteLine($"Template is not specified, skipping");
                                        processRow = false;
                                    }
                                    templatePath = value;
                                    break;
                                case "docgendate":
                                    indexOutputDocGenDate = j;
                                    break;
                                default:
                                    //just store all other params
                                    break;
                            }

                            //do not store unnamed values.
                            values.ElementAt(j - 1).value = value;
                            values.ElementAt(j - 1).type = type;
                        }

                    }
                }

                if(processRow)
                {
                    string outFileName = null;
                    DocumentProcessor run = new DocumentProcessor();
                    bool succeeded = run.Process(templatePath, values, out outFileName);
                    if(succeeded)
                    {
                        xlRange.Cells[i, indexOutputDoc].Value = outFileName;
                        xlRange.Cells[i, indexOutputDocGenDate].Value = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    }
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);

        }

        private void about_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.ShowDialog();
        }
    }
}
