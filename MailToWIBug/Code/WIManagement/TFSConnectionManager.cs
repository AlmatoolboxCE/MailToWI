using AlmaLogger;
using Microsoft.TeamFoundation.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MailToWIBug.Code
{
    public class TFSConnectionManager
    {
        protected string Url { get; set; }
        protected NetworkCredential Credentials;

        public TFSConnectionManager()
        {
            this.Url = ConfigurationManager.AppSettings["TFS.Url"];
            string user = ConfigurationManager.AppSettings["TFS.User"];
            string domain = ConfigurationManager.AppSettings["TFS.Domain"];
            string pwd = ConfigurationManager.AppSettings["TFS.Psswd"];
            this.Credentials = new NetworkCredential(user, pwd, domain);

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("Parametri di connessione a TFS:\n\tUrl: {0}\n\tUser: {1}\\{2}", this.Url, domain, user)));
        }

        public virtual TfsTeamProjectCollection GetConnection(string collection)
        {
            Uri tfsUri = new Uri(Url + "/" + collection);
            TfsTeamProjectCollection tfsServer = new TfsTeamProjectCollection(tfsUri, this.Credentials);
            return tfsServer;

        }

        public virtual TfsTeamProjectCollection GetConnection()
        {
            var collection = ConfigurationManager.AppSettings["TFS.Collection"];
            return GetConnection(collection);

        }
    }
}

