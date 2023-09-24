using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools
{
    public class JsonStorage
    {
        static string _fileName = null;
        public JsonStorage()
        {
        }

        private static void Initialize()
        {
            if (string.IsNullOrEmpty(_fileName))
                _fileName = CacheManager.GetFullPath("ConvertCV", "Settings");
        }
        public static void Write<T>(T objectToWrite, bool append = false) where T : new()
        {
            Initialize();

            var contentsToWriteToFile = JsonConvert.SerializeObject(objectToWrite, new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
            });
            using (var writer = new StreamWriter(_fileName, append))
            {
                writer.Write(contentsToWriteToFile);
            }
        }

        public static void Read<T>(T objectToWrite, bool append = false) where T : new()
        {
            Initialize();

            var contentsToWriteToFile = JsonConvert.SerializeObject(objectToWrite, new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
            });
            using (var writer = new StreamWriter(_fileName, append))
            {
                writer.Write(contentsToWriteToFile);
            }
        }
    }

    public class AppConfigurationSettings
    {
        public AppConfigurationSettings()
        {
            /* initialize the object if you want to output a new document
             * for use as a template or default settings possibly when 
             * an app is started.
             */
            if (AppSettings == null) { AppSettings = new AppSettings(); }
        }

        public AppSettings AppSettings { get; set; }
    }

    public class AppSettings
    {
        public bool DebugMode { get; set; } = false;
    }
}
