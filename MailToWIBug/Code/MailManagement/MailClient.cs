using AlmaLogger;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using MailToWIBug.Interface;

namespace MailToWIBug.Code
{
    public class MailClient<T>
    {

        /// <summary>
        /// Servizio Exchange.
        /// </summary>
        private ExchangeService service;

        protected bool EnableMailOnError { get; set; }
        public string AddressToMailOnError { get; set; }

        public MailClient()
        {
            string ewsUrl = ConfigurationManager.AppSettings["Exchange.Ews.Url"];
            string domain = ConfigurationManager.AppSettings["Exchange.Domain"];
            string username = ConfigurationManager.AppSettings["Exchange.User"];
            string password = ConfigurationManager.AppSettings["Exchange.Psswd"];

            this.EnableMailOnError = bool.Parse(ConfigurationManager.AppSettings["Mail.AlertOnError"]);
            this.AddressToMailOnError = ConfigurationManager.AppSettings["Mail.AlertOnErrorAddress"];
            this.AlertOnErrorTitle = ConfigurationManager.AppSettings["Mail.AlertOnErrorTitle"];


            if (string.IsNullOrWhiteSpace(AddressToMailOnError))
            {
                EnableMailOnError = false;
            }

            if (ewsUrl == null || domain == null || username == null || password == null)
            {
                ConfigurationErrorsException e = new ConfigurationErrorsException("Configurazione client di posta incompleta!");
                Logger.Fatal(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", string.Format("Errore nel caricamento dei parametri di connessione ad Exchange")));
                throw e;
            }

            string logging = string.Format(
                    "Connessione a \nUrl={0} \ndomain={1} \nusername={2} \npassword={3}",
                    ewsUrl,
                    domain,
                    username,
                    password);

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", logging));

            this.service = new ExchangeService(ExchangeVersion.Exchange2010);
            this.service.Url = new Uri(ewsUrl);
            this.service.Credentials = new NetworkCredential(username, password, domain);

            // Bind the Inbox folder to the service object
            Folder inbox = Folder.Bind(this.service, WellKnownFolderName.Inbox);

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", "Bind con la casella di posta effettuato correttamente."));
        }

        public ICollection<T> GetRequests()
        {
            ICollection<T> ret = new List<T>();

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", "Lettura mail..."));

            // Inizializziamo il filtro di ricerca per scaricare le nuove mail
            SearchFilter sf = new SearchFilter.SearchFilterCollection(
                            LogicalOperator.And,
                            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false) /*,
                            new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, "jarvis", ContainmentMode.Prefixed, ComparisonMode.IgnoreCase 

            )*/);

            ItemView view = new ItemView(20);

            // Lanciamo la query
            FindItemsResults<Item> findResults = this.service.FindItems(WellKnownFolderName.Inbox, sf, view);

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("Trovate {0} nuove mail", findResults.TotalCount)));

            if (findResults.TotalCount > 0)
            {

                List<Item> items = new List<Item>();

                foreach (EmailMessage emsg in findResults)
                {
                    Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("Elaboro {0}", emsg.Id)));
                    items.Add(emsg);
                    emsg.IsRead = true;
                    emsg.Update(ConflictResolutionMode.AutoResolve);
                }

                if (items.Count == 0) return ret;

                // si caricano le proprietà dei messaggi da leggere
                Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", "Caricamento proprieta' mail e creazione MailRequest"));

                PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.MimeContent);
                itempropertyset.RequestedBodyType = BodyType.HTML;
                this.service.LoadPropertiesForItems(items, itempropertyset);

                // si crea la Jarvis Request. NB: questa operazione va fatta dopo che le proprieta'
                // del messaggio sono state caricate
                foreach (EmailMessage mail in items)
                {


                    Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("MailRequest: {0}", mail.Subject)));
                    var ctr = typeof(T).GetConstructor(new Type[] { typeof(EmailMessage) });
                    if (ctr == null)
                    {
                        throw new Exception("Impossibile trovare un costruttore per soddisfare la richiesta");
                    }

                    //T request = new MailRequest(i.From.Address, i.ToRecipients.First().Address, i.Subject, i.Body);
                    T request = (T)ctr.Invoke(new object[] { mail });
                    ret.Add(request);
                }
            }

            return ret;
        }

        public void CleanEmail()
        {
            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", "Svuoto la caselal di posta"));

            //cerco tutte le mail già lette
            SearchFilter sf = new SearchFilter.SearchFilterCollection(
                            LogicalOperator.And,
                            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true) /*,
                            new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, "jarvis", ContainmentMode.Prefixed, ComparisonMode.IgnoreCase 

            )*/);

            ItemView view = new ItemView(20);

            // Lanciamo la query
            FindItemsResults<Item> findResults = this.service.FindItems(WellKnownFolderName.Inbox, sf, view);

            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("Trovate {0} nuove mail da cancellare", findResults.TotalCount)));

            if (findResults.TotalCount > 0)
            {
                foreach (var itm in findResults)
                {
                    itm.Delete(DeleteMode.HardDelete);
                }
            }
        }

        public void SendResponse(EmailMessage originalMessage, string body)
        {

            originalMessage.Load();
            Logger.Debug(new LogInfo(MethodBase.GetCurrentMethod(), "DEB", string.Format("Notifica evento al mittente {0}", originalMessage.Sender)));

            // Occorre "riagganciare" il messaggio alla casella
            // per farlo ci basta sapere l'id
            var propSet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.LastModifiedTime);

            // GetItem call to EWS.
            EmailMessage reply = EmailMessage.Bind(this.service, originalMessage.Id, propSet);

            reply.Reply(Mailify(body), true);

            //Console.WriteLine("Processato " + mailInput.EmailMessage.Subject);

        }

        private MessageBody Mailify(string msg)
        {
            MessageBody body = new MessageBody(BodyType.HTML, msg);
            return body;
        }

        public void SendLogErrorMail(string text)
        {
            if (!EnableMailOnError) return;

            var msg = new EmailMessage(this.service);
            msg.Subject = this.AlertOnErrorTitle;
            msg.Body = new MessageBody(text);
            msg.ToRecipients.Add(this.AddressToMailOnError);
            msg.Send();
        }

        public string AlertOnErrorTitle { get; set; }

        public void SendReplay(string to, string subject, byte[] attachment, string body)
        {
            var reply = new EmailMessage(this.service);
            reply.ToRecipients.Add(to);
            reply.Subject = subject;
            if (attachment != null)
            {
                reply.Attachments.AddFileAttachment("MessaggioOriginale.eml", attachment);
            }
            reply.Body = body;
            reply.Send();
        }
    }
}
