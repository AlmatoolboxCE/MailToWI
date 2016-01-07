using AlmaLogger;
using MailToWIBug.Interface;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MailToWIBug.Code.WIManagement
{
    public class WIAdder
    {
        public WIAdder()
        {
            this.ProjectName = ConfigurationManager.AppSettings["TFS.Project"];
            this.AreaPath = ConfigurationManager.AppSettings["TFS.AreaPath"];
            this.IterationPath = ConfigurationManager.AppSettings["TFS.IterationPath"];
            this.StepsToReproduceName = ConfigurationManager.AppSettings["TFS.StepsToReproduceName"];
            this.AssignedToName = ConfigurationManager.AppSettings["TFS.AssignedToName"];
            this.AssignedToValue = ConfigurationManager.AppSettings["TFS.AssignedToValue"];

        }

        public int AddNewWIFromMailBug(MessageRequest data)
        {
            var manager = new TFSConnectionManager();
            //var WIToAdd = GetNewWorkItem(Connection, ProjectName, "BUG", AreaPath, IterationPath, data.Subject, data.Body);

            using (var conn = manager.GetConnection())
            {
                WorkItemStore workItemStore = conn.GetService<WorkItemStore>();
                Project prj = workItemStore.Projects[ProjectName];
                WorkItemType workItemType = prj.WorkItemTypes["BUG"];


                var WIToAdd = new WorkItem(workItemType)
                {
                    Title = data.Subject,
                    Description = data.Body,
                    IterationPath = IterationPath,
                    AreaPath = AreaPath,
                };

                WIToAdd.Fields[this.StepsToReproduceName].Value = data.Body;
                if (!string.IsNullOrWhiteSpace(this.AssignedToValue))
                {
                    WIToAdd.Fields[this.AssignedToName].Value = this.AssignedToValue;
                }

                var filePath = TempFileManager.SaveFileAndGetName(data.FileFormat);
                try
                {
                    WIToAdd.Attachments.Add(new Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment(filePath));
                    WIToAdd.Save();
                    return WIToAdd.Id;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    TempFileManager.ClearFile(filePath);
                }
            }
        }

        public WorkItem GetNewWorkItem(TfsTeamProjectCollection conn, string projectName, string WIType, string areaPath, string iterationPath, string title, string description)
        {
            try
            {
                WorkItemStore workItemStore = conn.GetService<WorkItemStore>();
                Project prj = workItemStore.Projects[projectName];
                WorkItemType workItemType = prj.WorkItemTypes[WIType];


                var WIToAdd = new WorkItem(workItemType)
                {
                    Title = title,
                    Description = description,
                    IterationPath = iterationPath,
                    AreaPath = areaPath,
                };

                return WIToAdd;
            }
            catch (Exception ex)
            {
                Logger.Fatal(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", string.Format("Si è verificato un errore durante la creazione del work Item. Dettagli: {0}", ex.Message)));
                throw ex;
            }
        }

        public string ProjectName { get; set; }

        public string AreaPath { get; set; }

        public string IterationPath { get; set; }

        public string StepsToReproduceName { get; set; }

        public string AssignedToName { get; set; }

        public string AssignedToValue { get; set; }

        public WIDataForResponse GetWiData(string id)
        {
            var manager = new TFSConnectionManager();
            using (var conn = manager.GetConnection())
            {
                WorkItemStore workItemStore = conn.GetService<WorkItemStore>();
                WorkItemCollection queryResults = workItemStore.Query(
   "Select [State], [Title] " +
   "From WorkItems " +
   "Where [ID] = " + id);


                CredentialCache cc = new CredentialCache();
                cc.Add(
                    new Uri(ConfigurationManager.AppSettings["TFS.Url"]),
                    "NTLM",
                    new NetworkCredential(ConfigurationManager.AppSettings["TFS.User"],
                        ConfigurationManager.AppSettings["TFS.Psswd"],
                        ConfigurationManager.AppSettings["TFS.Domain"]));


                using (var wc = new WebClient()
                {
                    //Credentials = new NetworkCredential()
                    //{
                    //    Domain = ConfigurationManager.AppSettings["TFS.Domain"],
                    //    UserName = ConfigurationManager.AppSettings["TFS.User"],
                    //    Password = ConfigurationManager.AppSettings["TFS.Psswd"],
                    //},
                    Credentials = cc,
                })
                {
                    foreach (WorkItem wi in queryResults)
                    {
                        var atts = wi.Attachments;
                        foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.Attachment att in atts)
                        {
                            string tempName = null;
                            try
                            {
                                var data = wc.DownloadData(att.Uri.AbsoluteUri);
                                tempName = TempFileManager.SaveFileAndGetName(data);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception(string.Format("Impossibile ottenere l'attachment per il Work Item {0}.\nDettagli: {1}", id, ex.Message));
                            }
                            //mi serve solo il primo
                            return new WIDataForResponse()
                            {
                                AttachmentPath = tempName,
                                WIState = wi.State,
                                History = GetHistory(wi),
                            };
                        }
                    }
                }
            }

            throw new Exception(string.Format("Impossibile ottenere l'attachment per il Work Item {0}", id));
        }

        private string GetHistory(WorkItem wi)
        {
            //http://geekswithblogs.net/TarunArora/archive/2011/08/21/tfs-sdk-work-item-history-visualizer-using-tfs-api.aspx
            using (var dataTable = new DataTable())
            {

                foreach (Field field in wi.Fields)
                {
                    dataTable.Columns.Add(field.Name);
                }

                // Loop through the work item revisions
                foreach (Revision revision in wi.Revisions)
                {
                    // Get values for the work item fields for each revision
                    var row = dataTable.NewRow();
                    foreach (Field field in wi.Fields)
                    {
                        var value = revision.Fields[field.Name].Value == null ? string.Empty : revision.Fields[field.Name].Value.ToString();
                        row[field.Name] = value;
                    }
                    dataTable.Rows.Add(row);
                }

                // List of fields to ignore in comparison
                var visualize = new List<string>() { "Title", "State", "Rev", "Assigned To", };

                var result = new List<string>();

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    var currentRow = dataTable.Rows[i];


                    if (i + 1 < dataTable.Rows.Count)
                    {
                        var currentRowPlus1 = dataTable.Rows[i + 1];

                        //result.Add(String.Format("Comparing Revision {0} to {1} {2}", i, i + 1, Environment.NewLine));

                        bool title = false;

                        for (int j = 0; j < currentRow.ItemArray.Length; j++)
                        {
                            if (!title)
                            {
                                result.Add(
                                    String.Format(String.Format("Changed By '{0}' On '{1}'{2}", currentRow.Field<string>("Changed By"),
                                                                currentRow.Field<string>("Changed Date"), Environment.NewLine)));
                                title = true;
                            }

                            if (visualize.Contains(dataTable.Columns[j].ColumnName))
                            {
                                if (currentRow.ItemArray[j].ToString() != currentRowPlus1.ItemArray[j].ToString())
                                {
                                    result.Add(String.Format("[{0}]: '{1}' => '{2}' {3}", dataTable.Columns[j].ColumnName,
                                                              currentRow.ItemArray[j], currentRowPlus1.ItemArray[j],
                                                              Environment.NewLine));
                                }
                            }
                        }
                    }
                }
                return string.Join(string.Empty, result.ToArray());
            }
        }
    }
}
