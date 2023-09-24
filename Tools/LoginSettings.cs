using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Security;
using System.Configuration;
using System.IO;
using System.Security.Cryptography;

namespace Tools
{
    public class LoginSettings: SettingsBase
    {

        private LoginData _loginData = new LoginData();

        public LoginSettings(): base("login.xml")
        {
            Refresh();
        }

        public override void Refresh()
        {
            base.Refresh();

            SetParam(ref _loginData.UserName, "UserName");
            SetParam(ref _loginData.Password, "Password");
        }


        public override void Save()
        {
            var appSettings = _config.AppSettings;

            appSettings.Settings["UserName"].Value = _loginData.UserName;
            appSettings.Settings["Password"].Value = _loginData.Password;

            base.Save();
        }


        public String UserName
        {
            get
            {
                return SafeReturn(_loginData.UserName);
            }
            set
            {
                _loginData.UserName = StringProtector.Encrypt(value);
            }
        }
        public String Password
        {
            get
            {
                return SafeReturn(_loginData.Password);
            }
            set
            {
                _loginData.Password = StringProtector.Encrypt(value);
            }
        }
        /*public String ServerUrl
        {
            get
            {
                return SafeReturn(_loginData.ServerUrl);
            }
            set
            {
                _loginData.ServerUrl = StringProtector.Encrypt(value);
            }
        }

        public String ToolsHomeDir
        {
            get
            {
                return _loginData.ToolsHomeDir;
            }
            set
            {
                _loginData.ToolsHomeDir = value;
            }
        }*/

        private String SafeReturn(String encryptedString)
        {
            String result = "";
            try
            {
                result = StringProtector.Decrypt(encryptedString);
            }
            catch (CryptographicException ex)
            {
                result = "?????";
            }
            return result;

        }
    }
}
