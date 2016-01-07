using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using MailToWIBug.Code;
using MailToWIBug.Code.WIManagement;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Diagnostics;
using System.Threading;

using System.Linq;
using System.Linq.Expressions;
using System.Collections.Generic;
using System.IO;
namespace MailToWIBug.Test
{
    [TestClass]
    public class TFSConnectionTest
    {

        public class myFields
        {
            public string text { get; set; }
            public string value { get; set; }

            public override string ToString()
            {
                return string.Format("{0} - {1}", text, value);
            }
        }
        [TestMethod]
        public void TestMethod1()
        {
            var t = new Thread(() => { this.BodyThread(); });
            t.Start();
            while (true)
            {
                System.Threading.Thread.Sleep(100);
                if (t.ThreadState == System.Threading.ThreadState.Stopped || t.ThreadState == System.Threading.ThreadState.Aborted)
                {
                    break;
                }
            }
        }

        public void BodyThread()
        {

            var connManager = new TFSConnectionManager();

            using (var conn = connManager.GetConnection())
            {
                WorkItemStore workItemStore = conn.GetService<WorkItemStore>();
                foreach (Project itm in workItemStore.Projects)
                {
                    Debug.WriteLine(itm.Name);
                    foreach (WorkItemType workItemType in itm.WorkItemTypes)
                    {
                        Debug.WriteLine(workItemType.Name);
                        foreach (FieldDefinition f in workItemType.FieldDefinitions)
                        {
                            Debug.WriteLine(f.Name);
                        }
                    }
                }
                Assert.IsTrue(workItemStore.Projects.Count > 1);
            }
        }

        //[TestMethod]
        //public void TestMethod2()
        //{
        //    var connManager = new TFSConnectionManager();

        //    using (var conn = connManager.GetConnection("TEST2013"))
        //    {
        //        WorkItemStore workItemStore = conn.GetService<WorkItemStore>();
        //        foreach (Project itm in workItemStore.Projects)
        //        {
        //            Debug.WriteLine(itm.Name);
        //            foreach (WorkItemType workItemType in itm.WorkItemTypes)
        //            {
        //                Debug.WriteLine(workItemType.Name);
        //                foreach (FieldDefinition f in workItemType.FieldDefinitions)
        //                {
        //                    Debug.WriteLine(f.Name);
        //                }
        //            }
        //        }
        //        Assert.IsTrue(workItemStore.Projects.Count > 1);
        //    }
        //}

        [TestMethod]
        public void TestGetAttachment()
        {
            var tempPath = new MailToWIBug.Code.WIManagement.WIAdder().GetWiData("496").AttachmentPath;

            Assert.IsTrue(!string.IsNullOrWhiteSpace(tempPath));

        }

        [TestMethod]
        public void TestHistory()
        {
            var id = "17545";
            var result = new WIAdder().GetWiData(id);
            Assert.IsTrue(result.History.Length > 0);
            using(var wrt = new StreamWriter(string.Format("History_{0}.txt",id)))
            {
                wrt.Write(result.History);
            }
        }
    }
}
