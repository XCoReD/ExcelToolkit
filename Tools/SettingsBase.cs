using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class SettingsBase
    {
        protected System.Configuration.Configuration _config;
        private String _settingsFileName;

        internal class StringProtector
        {
            //http://eddiejackson.net/wp/?p=13434

            // set permutations
            public const String strPermutation = "ouiveyxaqtd";
            public const Int32 bytePermutation1 = 0x19;
            public const Int32 bytePermutation2 = 0x59;
            public const Int32 bytePermutation3 = 0x17;
            public const Int32 bytePermutation4 = 0x41;

            // encoding
            public static string Encrypt(string strData)
            {

                return Convert.ToBase64String(Encrypt(Encoding.UTF8.GetBytes(strData)));
                // reference https://msdn.microsoft.com/en-us/library/ds4kkd55(v=vs.110).aspx

            }


            // decoding
            public static string Decrypt(string strData)
            {
                return Encoding.UTF8.GetString(Decrypt(Convert.FromBase64String(strData)));
                // reference https://msdn.microsoft.com/en-us/library/system.convert.frombase64string(v=vs.110).aspx

            }

            // encrypt
            public static byte[] Encrypt(byte[] strData)
            {
                PasswordDeriveBytes passbytes =
                new PasswordDeriveBytes(strPermutation,
                new byte[] { bytePermutation1,
                         bytePermutation2,
                         bytePermutation3,
                         bytePermutation4
                });

                MemoryStream memstream = new MemoryStream();
                Aes aes = new AesManaged();
                aes.Key = passbytes.GetBytes(aes.KeySize / 8);
                aes.IV = passbytes.GetBytes(aes.BlockSize / 8);

                CryptoStream cryptostream = new CryptoStream(memstream,
                aes.CreateEncryptor(), CryptoStreamMode.Write);
                cryptostream.Write(strData, 0, strData.Length);
                cryptostream.Close();
                return memstream.ToArray();
            }

            // decrypt
            public static byte[] Decrypt(byte[] strData)
            {
                PasswordDeriveBytes passbytes =
                new PasswordDeriveBytes(strPermutation,
                new byte[] { bytePermutation1,
                         bytePermutation2,
                         bytePermutation3,
                         bytePermutation4
                });

                MemoryStream memstream = new MemoryStream();
                Aes aes = new AesManaged();
                aes.Key = passbytes.GetBytes(aes.KeySize / 8);
                aes.IV = passbytes.GetBytes(aes.BlockSize / 8);

                CryptoStream cryptostream = new CryptoStream(memstream,
                aes.CreateDecryptor(), CryptoStreamMode.Write);
                cryptostream.Write(strData, 0, strData.Length);
                cryptostream.Close();
                return memstream.ToArray();
            }
            // reference
            // https://msdn.microsoft.com/en-us/library/system.security.cryptography(v=vs.110).aspx
            // https://msdn.microsoft.com/en-us/library/system.security.cryptography.cryptostream%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396
            // https://msdn.microsoft.com/en-us/library/system.security.cryptography.rfc2898derivebytes(v=vs.110).aspx
            // https://msdn.microsoft.com/en-us/library/system.security.cryptography.aesmanaged%28v=vs.110%29.aspx?f=255&MSPPError=-2147217396

        }

        public SettingsBase(string fileName, String product = "Global")
        {
            String folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DT");
            Directory.CreateDirectory(folder);
            folder = Path.Combine(folder, product);
            Directory.CreateDirectory(folder);
            _settingsFileName = folder + "\\" + fileName;

            Refresh();
            //if (File.Exists(settingsFile))
            //    ConfigurationManager.RefreshSection("appSettings");

        }

        public String SettingsFileName
        {
            get
            {
                return _settingsFileName;
            }
        }

        virtual public void Save()
        {
             _config.Save(ConfigurationSaveMode.Modified);
        }

        virtual public void Refresh()
        {
            ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
            configMap.ExeConfigFilename = _settingsFileName;

            _config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
            //ConfigurationManager.RefreshSection("appSettings");
        }

        protected void SetParam(ref String destValue, String name)
        {
            var param = _config.AppSettings.Settings[name];
            if (param == null)
                _config.AppSettings.Settings.Add(name, "");
            else
                destValue = param.Value;
        }

        protected void SetParam(ref bool destValue, String name, bool defValue = false)
        {
            destValue = defValue;
            var param = _config.AppSettings.Settings[name];
            if (param == null)
                _config.AppSettings.Settings.Add(name, "");
            else
                destValue = Boolean.Parse(param.Value);
        }

    }
}
