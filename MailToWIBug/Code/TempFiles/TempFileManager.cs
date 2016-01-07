using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToWIBug.Code
{
    class TempFileManager
    {
        public static int counter = 0;
        public static object semaphore = new object();
        public static string TempDir = "TEMP";

        public static string SaveFileAndGetName(byte[] fileBytes)
        {
            var filePath = string.Empty;
            lock (semaphore)
            {
                var dirPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, TempDir);
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }
                filePath = string.Format("{0}\\testFileMail{1}.eml", dirPath, counter);
                counter++;
            }
            using (var wrt = File.Create(filePath))
            {
                wrt.Write(fileBytes, 0, fileBytes.Length);
            }

            return filePath;
        }

        public static void ClearFile(string filePath)
        {
            File.Delete(filePath);
        }
    }
}
