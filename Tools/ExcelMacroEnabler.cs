using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class ExcelMacroEnabler
    {
        public static string GetSafePath()
        {
            string safeLocationPlace = "\\Microsoft\\Excel\\XLSTART";
            string safeLocationTemplate = $"%APPDATA%{safeLocationPlace}";
            //https://stackoverflow.com/questions/3266675/how-to-detect-installed-version-of-ms-office
            string[] versions = new string[] { "16.0", "15.0", "14.0", "12.0", "11.0" };
            RegistryHive hive = RegistryHive.CurrentUser;
            using (var hiveKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
            {
                foreach(var version in versions)
                {
                    string path = $"Software\\Microsoft\\Office\\{version}\\Excel\\Security\\Trusted Locations";
                    using (var key = hiveKey.OpenSubKey(path))
                    {
                        string lastName = null;
                        var subKeys = key.GetSubKeyNames();
                        foreach (var subkey in subKeys)
                        {
                            using (var location = hiveKey.OpenSubKey(path + "\\" + subkey))
                            {
                                lastName = location.Name;
                                //Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\Location1
                                RegistryValueKind rvk = location.GetValueKind("Path");
                                if (rvk == RegistryValueKind.ExpandString)
                                {
                                    string locationPath = (string)location.GetValue("Path");
                                    if (locationPath.IndexOf(safeLocationPlace) > 0)
                                        return locationPath;
                                }
                            }
                        }
                        if(!string.IsNullOrEmpty(lastName))
                        {
                            //can add another one
                            const string prefix = "Location";
                            int pos = lastName.LastIndexOf(prefix);
                            Debug.Assert(pos >= 0);
                            if(pos >= 0)
                            {
                                lastName = lastName.Substring(pos + prefix.Length);
                                int index = 0;
                                if(Int32.TryParse(lastName, out index))
                                {
                                    lastName = $"{prefix}{index + 1}";
                                    try
                                    {
                                        var createdKey = key.CreateSubKey(lastName);
                                        createdKey.SetValue("Description", "4", RegistryValueKind.String);
                                        createdKey.SetValue("Path", safeLocationTemplate, RegistryValueKind.ExpandString);

                                        string result = Environment.GetEnvironmentVariable("%APPDATA%") + safeLocationPlace;
                                        return result;
                                    }
                                    catch(UnauthorizedAccessException ex)
                                    {
                                        //TODO: handle
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }

    }
}
