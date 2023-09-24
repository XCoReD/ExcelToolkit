using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.AccessControl;

namespace ExcelToolkit
{
    class Settings
    {
        //C:\Users\dtrus\AppData\Roaming\Microsoft\AddIns

        Microsoft.Office.Interop.Excel.Application _app;
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public static string AddinName
        {
            get
            {
                return "DT ExcelFunctions";
            }
        }

        private string XllName
        {
            get
            {
                String xllName;
                if (System.Runtime.InteropServices.Marshal.SizeOf(_app.HinstancePtr) == 8)
                {
                    // excel 64-bit
                    //xllName = "ExcelLib-AddIn64-packed.xll";
                    //xllName = "ExcelLib-AddIn64.xll";
                    xllName = "ExcelFunctions-AddIn64-packed.xll";
                }
                else
                {
                    // excel 32-bit
                    //xllName = "ExcelLib-AddIn-packed.xll";
                    //xllName = "ExcelLib-AddIn.xll";
                    xllName = "ExcelFunctions-AddIn-packed.xll";
                }
                return xllName;
            }
        }

        public void UnRegisterFunctions()
        {
            _app = null;
        }

        bool GetInstalledState()
        {
            var name = XllName.ToUpper();
            foreach (Microsoft.Office.Interop.Excel.AddIn addin in _app.AddIns2)
            {
                string d = addin.Name;
                Debug.WriteLine(d);
                if (d.ToUpper() == name)
                {
                    return addin.Installed;
                }
            }
            return false;
        }
        bool SetInstalledState(bool state)
        {
            var name = XllName.ToUpper();
            Debug.WriteLine("Excel addins:");
            foreach (Microsoft.Office.Interop.Excel.AddIn addin in _app.AddIns2)
            {
                string d = addin.Name;
                Debug.WriteLine(d);
                if (d.ToUpper() == name)
                {
                    if (addin.Installed == state)
                    {
                        Debug.WriteLine($"Add-in in {state} state, no need to change");
                        return true;
                    }
                    try
                    {
                        addin.Installed = state;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Exception caught when turning off add-in: " + ex.Message);
                    }
                    break;
                }
            }
            return false;
        }

        void SetAccessControl(string path)
        {
            var directoryInfo = Directory.CreateDirectory(path);
            bool modified;
            var directorySecurity = directoryInfo.GetAccessControl();
            SecurityIdentifier securityIdentifier = new SecurityIdentifier
    (WellKnownSidType.BuiltinUsersSid, null);

            var rule = new FileSystemAccessRule(
                securityIdentifier,
                FileSystemRights.Write |
                FileSystemRights.ReadAndExecute |
                FileSystemRights.Modify |
                FileSystemRights.Delete,
                InheritanceFlags.ContainerInherit |
                InheritanceFlags.ObjectInherit,
                PropagationFlags.InheritOnly,
                AccessControlType.Allow);
            directorySecurity.ModifyAccessRule(AccessControlModification.Add, rule, out modified);
            directoryInfo.SetAccessControl(directorySecurity);
        }
        public bool RegisterFunctions()
        {
            string localXllName = GetLocalXllName();
            string installedXllAddinPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Microsoft",
                "AddIns");
            string installedXllName = Path.Combine(installedXllAddinPath, XllName);
            if(File.Exists(installedXllName))
            {
                var dtExisting = File.GetLastWriteTime(installedXllName);
                var dtLocal = File.GetLastWriteTime(localXllName);
                if (dtExisting >= dtLocal)
                {
                    Console.WriteLine("Xll is installed and up to date");
                    return true;
                }
                var diff = dtLocal - dtExisting;
                if (diff.TotalMinutes < 1)
                    return true;

                bool installed = GetInstalledState();
                if(installed)
                {
                    SetInstalledState(false);
                    MessageBox.Show($"You got update of Excel DT plugin.\n\nPlease restart Excel to finish.\n", "ExcelToolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                try
                {
                    //SetAccessControl(installedXllAddinPath);

                    File.Delete(installedXllName);
                    File.Copy(localXllName, installedXllName);
                }
                catch(Exception ex)
                {
                    //Debug.Assert(false, ex.Message);
                    return false;
                }
            }
            else
            {
                try
                {
                    File.Copy(localXllName, installedXllName);
                }
                catch (Exception ex)
                {
                    //Debug.Assert(false, ex.Message);
                }

                bool result = RegisterXll(_app, XllName);
                if (!result)
                {
                    MessageBox.Show($"Excel plugin registration failed.\n\nPlease do it manually, following guide at Confluence.\n", "ExcelToolkit", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            /*var fullPath = Path.Combine(AssemblyDirectory, XllName);
            AddIn addin = null;
            try
            {
                addin = _app.AddIns2.Add(fullPath, true); //true
                Debug.Assert(addin != null);
                return true;
            }
            catch (COMException ex)
            {
                //MessageBox.Show($"Exception when adding addin: {ex.Message}, {ex.Source}, {ex.StackTrace}");
                Debug.Assert(false, ex.Message);
            }*/

            SetInstalledState(true);

            return true;
        }

        private string GetLocalXllName()
        {
            string fullName = null;
#if DEBUG
            string location = AssemblyDirectory;
#else
            string location = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
#endif
            fullName = Path.Combine(location, XllName);

            //MessageBox.Show(fullName);

#if !DEBUG
            if (!File.Exists(fullName))
#endif
            {
                //looks this is the first time we run - need to extract from resources
                string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                using (var rc = Assembly.GetExecutingAssembly().GetManifestResourceStream($"{assemblyName}.{XllName}"))
                {
                    using (var file = new FileStream(fullName, FileMode.Create, FileAccess.Write))
                    {
                        rc.CopyTo(file);
                    }
                    var dt = BuildDate.GetBuildDateTime(System.Reflection.Assembly.GetExecutingAssembly());
                    File.SetCreationTime(fullName, dt);
                    File.SetLastWriteTime(fullName, dt);
                }
            }
            return fullName;

        }
        private static bool RegisterXll(Microsoft.Office.Interop.Excel.Application app, string name)
        {
            String dir = Directory.GetCurrentDirectory();
#if !DEBUG
            string location = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
#else
            string location = AssemblyDirectory;
#endif
            Directory.SetCurrentDirectory(location);

            bool result = app.RegisterXLL(name);
            //MessageBox.Show($"Name: {name}, Registered: {result}");
            Debug.Assert(result == true);
            Directory.SetCurrentDirectory(dir);

            return result;
        }

        public Settings(Microsoft.Office.Interop.Excel.Application app)
        {
            _app = app;
        }

    }
}
