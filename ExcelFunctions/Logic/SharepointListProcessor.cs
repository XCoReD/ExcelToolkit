using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;
using System.Net;

namespace ExcelFunctions
{
    class SharepointListProcessor
    {
        ILog _log;
        Worksheet _ws;
        const string ActionColumnName = "<<SP_Action>>";
        int _actionColumnId = 0;
        const string ActionResultColumnName = "<<SP_Action_Result>>";
        int _actionResultColumnId = 0;

        int _modifiedColumnId = 0;
        int _guidColumnId = 0;

        string _listInfo;
        ClientContext _clientContext;
        Microsoft.SharePoint.Client.List _list;
        string[] _visibleColumnsInfo;
        public SharepointListProcessor(Worksheet ws, ILog log, string listInfo)
        {
            _ws = ws;
            _log = log;
            _listInfo = listInfo;
        }

        public static bool Process(Microsoft.Office.Interop.Excel.Application app, ILog log)
        {
            bool result = false;
            Workbook wb = app.ActiveWorkbook;
            if(wb == null)
            {
                log.Error("No open workbook");
                return false;

            }
            if (wb.Connections == null)
            {
                log.Error("No external data connections found");
                return false;
            }

            string listInfo = null;
            foreach (WorkbookConnection connection in wb.Connections)
            {
                /*if (connection.Type == XlConnectionType.xlConnectionTypeDATAFEED)
                {
                    DataFeedConnection dataFeedCur = connection.DataFeedConnection;
                    probablyConnectionFound = true;
                }
                else*/ if(connection.Type == XlConnectionType.xlConnectionTypeOLEDB)
                {
                    OLEDBConnection conn = connection.OLEDBConnection;
                    if(conn.CommandType == XlCmdType.xlCmdList)
                    {
                        //sharepoint list
                        Debug.Assert(listInfo == null);
                        listInfo = conn.CommandText;
                    }
                }
            }

            if(string.IsNullOrEmpty(listInfo))
            {
                log.Error("No connections to Sharepoint found");
                return false;
            }

            log.Info("Connection info:" + listInfo);

            Worksheet ws = app.ActiveSheet;
            if(ws == null)
            {
                log.Error("No active worksheet found");
                return false;
            }

            SharepointListProcessor instance = new SharepointListProcessor(ws, log, listInfo);
            return instance.Process();

            /*int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            log.BeginSession("Fill Template");

            List<Param> arguments = new List<Param>(colCount);*/

            /*for (int i = 1; i <= rowCount; i++)
            {
                List<Param> values = i == 1 ? arguments : arguments.GetRange(0, arguments.Count);
                bool processRow = i > 1;
                int indexOutputDoc = 0;
                int indexOutputDocGenDate = 0;
                string templatePath = null;

                for (int j = 1; j <= colCount; j++)
                {
                    if (i == 1)
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

                    }
                }
            }*/

            //return result;
        }

        public bool Process()
        {
            if (!HasSharepointListContent())
            {
                _log.Error("The spreadsheet has no rows. Please fetch data from Sharepoint list first");
                return false;

            }

            //check if there is a column "Action"
            _actionColumnId = FindColumn(ActionColumnName);
            if (_actionColumnId == 0)
            {
                _log.Error("No action column found, adding..");
                _actionColumnId = AddColumn(ActionColumnName);

                _log.Error("Filling in supposed action for each cell..");
                PrognozeActions();

                string msg = "The current spreadsheet had no planned action column. This column has been added, and the suggested operations were added.\n\nPlease review and run sync again";
                MessageBox.Show(msg, "Excel Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return true;
            }

            _actionResultColumnId = FindColumn(ActionResultColumnName);
            if (_actionResultColumnId == 0)
            {
                _log.Error("No action result column found, added.");
                _actionResultColumnId = AddColumn(ActionResultColumnName);
            }

            if(!Connect())
            {
                return false;
            }

            FillVisibleColumnsInfo();

            if (!ProcessActions())
            {
                return false;
            }

            return true;


        }

        void FillVisibleColumnsInfo()
        {
            int i = 0;
            _visibleColumnsInfo = new string[_ws.Columns.Count];
            foreach(Range column in _ws.Columns)
            {
                ++i;
                if(!column.Hidden)
                {
                    if (_ws.Cells[1, i].Value2 != null)
                    {
                        string name = _ws.Cells[1, i].Value2.ToString();
                        if(!IsSpecialColumn(name))
                            _visibleColumnsInfo[i] = name;
                    }
                }
            }
        }

        bool IsSpecialColumn(string columnName)
        {
            if(
                columnName == ActionColumnName ||
                columnName == ActionResultColumnName ||
                columnName == "Modified" ||
                columnName == "Created" ||
                columnName == "AuthorId" ||
                columnName == "EditorId" ||
                columnName == "OData__UIVersionString" ||
                columnName == "Attachments" ||
                columnName == "GUID")
            {
                return true;
            }
            return false;
        }
        bool Connect()
        {
            string url = GetInnerTagContent(_listInfo, "LISTWEB");
            Debug.Assert(false);
            return false;

            string siteUrl = "https://yoursharepointportal.com";

            _clientContext = new ClientContext(siteUrl);
            _clientContext.Credentials = CredentialCache.DefaultCredentials;

            //cut this: <LISTNAME>{5D417A25-F298-4F03-BE0A-F675222CB7EE}</LISTNAME>
            Guid guid = new Guid(GetInnerTagContent(_listInfo, "LISTNAME"));

            _list = _clientContext.Web.Lists.GetById(guid);
            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = _list.GetItems(camlQuery);


            try
            {
                _clientContext.Load(
                    collListItem,
                    items => items.Take(5).Include(
                    item => item["GUID"]));

                _clientContext.ExecuteQuery();


                /*_clientContext.Load(_list);
                _clientContext.ExecuteQuery();*/
            }
            catch (System.Net.WebException ex)
            {
                _log.Error("Failed to connect to sharepoint list", ex);
                return false;
            }
            catch (ServerException ex)
            {
                _log.Error("Failed to connect to sharepoint list", ex);
                return false;
            }


            Debug.Assert(_list != null);

            return true;
        }

        string GetInnerTagContent(string text, string tagName)
        {
            string tagNameBegin = "<" + tagName + ">";
            string tagNameEnd = "</" + tagName + ">";
            string result = null;
            int i = text.IndexOf(tagNameBegin);
            if (i != -1)
            {
                int iEnd = text.IndexOf(tagNameEnd, i + tagNameBegin.Length);
                if (iEnd != -1)
                {
                    result = text.Substring(i + tagNameBegin.Length, iEnd - i - tagNameBegin.Length);
                }
            }
            return result;
        }
        bool HasSharepointListContent()
        {
            Range xlRange = _ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            if(rowCount == 0)
            {
                _log.Error("No rows in the spreadsheet");
                return false;
            }

            if (colCount == 0)
            {
                _log.Error("No columns in the spreadsheet");
                return false;
            }

            _modifiedColumnId = FindColumn("Modified");
            _guidColumnId = FindColumn("GUID");
            if (_modifiedColumnId == 0 ||
                FindColumn("Created") == 0 ||
                FindColumn("AuthorId") == 0 ||
                FindColumn("EditorId") == 0 ||
                FindColumn("OData__UIVersionString") == 0 ||
                FindColumn("Attachments") == 0 ||
                _guidColumnId == 0)
            {
                _log.Error("No mandatory columns for Sharepoint list found");
                return false;
            }

            return true;
        }

        int FindColumn(string name)
        {
            Range xlRange = _ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= colCount; ++i)
            {
                Range range = xlRange.Cells[1, i];
                if (range != null && range.Value2 != null)
                {
                    string value = range.Value2.ToString();
                    if (name == value)
                        return i;
                }
            }
            return 0;
        }

        int AddColumn(string name)
        {
            //_ws.UsedRange.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlToRight);
            int addedColumn = _ws.UsedRange.Columns.Count + 1;
            _ws.Columns[addedColumn].Insert(Microsoft.Office.Interop.Excel.XlDirection.xlToRight);
            _ws.UsedRange.Cells[1, addedColumn].Value2 = name;
            return addedColumn;
        }

        bool PrognozeActions()
        {
            //3 actions - C,U,D
            Range xlRange = _ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                Range range = xlRange.Cells[i, _modifiedColumnId];
                string verb = null;
                if (range == null || range.Value2 == null)
                {
                    verb = "C";
                }
                else
                {
                    verb = "U";
                }
                if(verb != null)
                {
                    xlRange.Cells[i, _actionColumnId].Value2 = verb;
                }
            }
            return true;
        }

        bool ProcessActions()
        {
            Range xlRange = _ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            bool result = true;

            for (int i = 2; i <= rowCount; i++)
            {
                Range range = xlRange.Cells[i, _actionColumnId];
                if (range != null && range.Value2 != null)
                {
                    string operation = range.Value2.ToString();
                    if (!string.IsNullOrEmpty(operation))
                    {
                        switch (operation.ToLower().Trim())
                        {
                            case "c":
                            case "create":
                            case "new":
                                if (!UpdateRowResult(i, Create(i)))
                                    result = false;
                                break;
                            case "u":
                            case "update":
                                if(!UpdateRowResult(i, Update(i)))
                                    result = false;
                                break;
                            case "d":
                            case "delete":
                                if (!UpdateRowResult(i, Delete(i)))
                                    result = false;
                                break;
                            default:
                                Debug.Assert(false);
                                break;
                        }
                    }
                }
            }
            return result;
        }

        bool UpdateRowResult(int rowNumber, bool result)
        {
            Range cell = _ws.UsedRange.Cells[rowNumber, _actionResultColumnId];
            cell.Value2 = result.ToString();
            return result;
        }

        bool Create(int rowNumber)
        {
            Range cell = _ws.UsedRange.Cells[rowNumber, 2];
            if(cell.Value2 != null)
            {

            }

            /*ListItem oListItem = _list.Items.GetById(3);

            oListItem["Title"] = "My Updated Title.";

            oListItem.Update();

            clientContext.ExecuteQuery();*/

            return true;
        }

        bool Update(int rowNumber)
        {
            Range xlRange = _ws.UsedRange;
            int colCount = xlRange.Columns.Count;

            Guid guid = new Guid(xlRange.Cells[rowNumber, _guidColumnId].Value2.ToString());
            Debug.Assert(!guid.Equals(Guid.Empty));
            if(guid.Equals(Guid.Empty))
            {
                _log.Error($"Updating row {rowNumber} is not possible: guid is empty");
                return false;
            }

            ListItem oListItem = _list.GetItemById(guid.ToString());
            Debug.Assert(oListItem != null);

            string logInfo = "Updating " + guid.ToString() + ": ";
            int updatedValues = 0;
            for (int i = 1; i <= colCount; i++)
            {
                if(_visibleColumnsInfo[i] != null)
                {
                    Range range = xlRange.Cells[rowNumber, i];
                    if (range != null && range.Value2 != null)
                    {
                        string stringValue = range.Value2.ToString();
                        if(!string.IsNullOrEmpty(stringValue))
                        {
                            logInfo += _visibleColumnsInfo[i] + "=" + stringValue + ";";
                            oListItem[_visibleColumnsInfo[i]] = stringValue;
                            ++updatedValues;
                        }
                    }
                }
            }

            _log.Info(logInfo);

            if(updatedValues != 0)
            {
                try
                {
                    oListItem.Update();
                    _clientContext.ExecuteQuery();
                }
                catch(Exception ex)
                {
                    _log.Error(logInfo + " - failed", ex);
                    return false;
                }
            }
            return true;
        }

        bool Delete(int rowNumber)
        {
            return true;
        }
    }

}
