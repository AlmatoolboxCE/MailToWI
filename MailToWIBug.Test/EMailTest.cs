using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MailToWIBug.Code;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;

namespace MailToWIBug.Test
{
    [TestClass]
    public class EMailTest
    {
        [TestMethod]
        public void AttachementsTest()
        {
            var coda = new Queue<MailRequest>();
            var obs = new MailObserver<MailRequest>(coda);
            var t = new Thread(() => { obs.MailClientLoop(); });
            int counter = 0;
            t.Start();

            var fileName = string.Empty;

            lock (coda)
            {
                // Console.WriteLine("Jarvis waits ...");
                if (coda.Count == 0)
                {
                    Monitor.Pulse(coda);
                    Monitor.Wait(coda);
                }

                // svuoto la coda
                while (coda.Count != 0)
                {
                    var request = coda.Dequeue();
                    fileName = string.Format("{0}\\testFileMail{1}.eml", Directory.GetCurrentDirectory(), counter);
                    using (var wrt = File.Create(fileName))
                    {
                        wrt.Write(request.FileFormat, 0, request.FileFormat.Length);
                    }
                }

                Monitor.PulseAll(coda);
                t.Abort();
            }

            Assert.IsTrue(File.Exists(fileName));
        }


        [TestMethod]
        public void CLeanEmail()
        {
            var coda = new Queue<MailRequest>();
            var obs = new MailObserver<MailRequest>(coda);
            obs.Client.CleanEmail();
            Assert.IsTrue(true);
        }


        [TestMethod]
        public void TestErroreMail()
        {
            var toTest = new MailToWIBugService<MailRequest>();
            //using (var rdr = new StreamReader(string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "testFileMail23.eml")))
            //using (var rdr = new StreamReader(string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "testMailStrana.txt")))
            using (var rdr = new StreamReader(string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "testFileMail8.eml")))
            {
                var strs = rdr.ReadToEnd();
                var rows = strs.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                rows = toTest.RemoveCodedRows(rows);

                var subj = toTest.GetMailSubject(rows, "0");
                Assert.IsTrue(subj.Length > 0);
            }
        }
    }
}
