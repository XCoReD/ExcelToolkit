using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace Tools
{
    public class CacheManager
    {
        public static string GetFullPath(string appName, string fileName, bool makeFolders = true)
        {
            String folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DT");
            if (makeFolders)
                Directory.CreateDirectory(folder);

            folder = Path.Combine(folder, appName);
            if (makeFolders)
                Directory.CreateDirectory(folder);

            folder = Path.Combine(folder, "Cache");
            Directory.CreateDirectory(folder);
            if (makeFolders)
                Directory.CreateDirectory(folder);

            return Path.Combine(folder, fileName);
        }

        public static void ResetCache(string appName, string fileName)
        {
            String name = GetFullPath(appName, fileName, false);
            if (File.Exists(name))
                File.Delete(name);
        }

    }
    public class ObjectSerializer<T>
    {
        protected IFormatter iformatter;
        protected int _cacheAliveDays;
        protected String _fileName;
        //dic = new ConcurrentDictionary<K, CacheData<V>>(dic.Where(x => DateTime.Now.Subtract(x.Value.CreationDate).TotalMilliseconds<_dormantCacheExpire));

        public ObjectSerializer(string appName, string fileName, int cacheAliveDays = 180)
        {
            _cacheAliveDays = cacheAliveDays;
            this.iformatter = new BinaryFormatter();

            _fileName = CacheManager.GetFullPath(appName, fileName);
        }

        public T GetSerializedObject(bool readIfObsolete = false)
        {
            if (Exists())
            {
                if(readIfObsolete || !Obsolete())
                {
                    Stream inStream = new FileStream(
                    _fileName,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read);

                    T obj = (T)this.iformatter.Deserialize(inStream);
                    inStream.Close();

                    return obj;
                }
            }

            return default(T);

        }

        public bool Exists()
        {
            return File.Exists(_fileName);
        }

        public string GetFileName()
        {
            return _fileName;
        }
        public bool Obsolete()
        {
            if (DateTime.Now.Subtract(File.GetLastWriteTime(_fileName)).TotalDays > _cacheAliveDays)
                return true;

            return false;
        }

        public void SaveSerializedObject(T obj)
        {
            Stream outStream = new FileStream(
            _fileName,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None);
            this.iformatter.Serialize(outStream, obj);

            outStream.Close();
        }
    }
}
