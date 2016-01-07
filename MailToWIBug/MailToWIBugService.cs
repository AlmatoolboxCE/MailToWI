//----------------------------------------------------------------------------------------------
// <copyright file="MailToWIBugService.cs" company="AlmavivA" author="a.delucia@almaviva.it Andrea De Lucia">
// Copyright (c) AlmavivA.  All rights reserved.
// </copyright>
//----------------------------------------------------------------------------------------------
namespace MailToWIBug
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Configuration;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.ServiceProcess;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;

    using AlmaLogger;

    using MailToWIBug.Code;
    using MailToWIBug.Code.WIManagement;
    using MailToWIBug.Interface;

    /// <summary>
    /// Servizio di ascolto ed elaborazione delle mail
    /// </summary>    
    public partial class MailToWIBugService<T> : ServiceBase
        where T : MessageRequest
    {
        /// <summary>
        /// contiene le stringhe che determinano l'eliminazione del messaggio
        /// </summary>        
        private string[] automaticResponseString = ConfigurationManager.AppSettings["Mail.IgnoreStrings"].Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

        /// <summary>
        /// 
        /// </summary>
        private string regexMittenteMail = @"from:[\s\w]*<*(?<data>(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9])))>*";

        /// <summary>
        /// 
        /// </summary>
        private string regexOggettoMail = @"subject:\s*(?<data>.*)";

        internal string base64Pattern = @"=\?utf-8\?B\?(?<data>.*)\?=";

        internal string quotedPattern = @"=\?utf-8\?Q\?(?<data>.*)\?=";

        /// <summary>
        /// contiene il pattern di ricerca dell'ID del Work Item
        /// </summary>
        private string regexPatternToWIId = @"Work\sitem\sChanged:\sBug\s(?<data>[0-9]+)\s-";

        /// <summary>
        /// contiene l'elenco degli indirizzi da cui arrivano le risposte di TFS
        /// </summary>        
        private string[] replyAddresses = ConfigurationManager.AppSettings["Mail.Response.TFSAddress"].ToLower().Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

        /// <summary>
        /// costruttore di default
        /// </summary>
        public MailToWIBugService()
        {
            this.InitializeComponent();
            this.Workers = new Thread[1];
            this.MailClients = new Thread[1];
            this.MailQueue = new Queue<T>();
            this.TheMailObserver = new MailObserver<T>(this.MailQueue);
            this.WIManager = new WIAdder();
        }

        /// <summary>
        /// Gestore dei WI
        /// </summary>
        public WIAdder WIManager
        {
            get;
            set;
        }

        /// <summary>
        /// thread di ascolto sulle e-mail
        /// </summary>
        protected Thread[] MailClients
        {
            get;
            set;
        }

        /// <summary>
        /// coda di mail in arrivo
        /// </summary>
        protected Queue<T> MailQueue
        {
            get;
            set;
        }

        /// <summary>
        /// sistema di ascolto delle mail
        /// </summary>
        protected MailObserver<T> TheMailObserver
        {
            get;
            set;
        }

        /// <summary>
        /// thread di elaborazione
        /// </summary>
        protected Thread[] Workers
        {
            get;
            set;
        }

        /// <summary>
        /// inizializzazione del servizio
        /// </summary>
        public virtual void InitApplication()
        {
            this.InitLogger();
            try
            {
                // init service workers
                for (int i = 0; i < this.Workers.Length; i++)
                {
                    var t = this.Workers[i];
                    if (t != null && t.ThreadState == System.Threading.ThreadState.Running)
                    {
                        t.Abort();
                    }

                    t = new Thread(() => { this.OLoop(); });
                    t.Start();
                    this.Workers[i] = t;
                }

                // init mail daemons
                for (int i = 0; i < this.MailClients.Length; i++)
                {
                    var t = this.MailClients[i];
                    if (t != null && t.ThreadState == System.Threading.ThreadState.Running)
                    {
                        t.Abort();
                    }

                    t = new Thread(() => { this.TheMailObserver.MailClientLoop(); });
                    t.Start();
                    this.MailClients[i] = t;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "ERR", string.Format("Errore all'avvio del servizio. Dettagli: {0}", ex.Message)));
            }
        }

        /// <summary>
        /// ciclo di elaborazione
        /// </summary>
        public virtual void OLoop()
        {
            while (true)
            {
                lock (this.MailQueue)
                {
                    // Console.WriteLine("Jarvis waits ...");
                    if (this.MailQueue.Count == 0)
                    {
                        Monitor.Pulse(this.MailQueue);
                        Monitor.Wait(this.MailQueue);
                    }

                    // svuoto la coda
                    while (this.MailQueue.Count != 0)
                    {
                        try
                        {
                            var request = this.MailQueue.Dequeue();
                            switch (this.CheckElaborateThisMail(request))
                            {
                                case MailAction.Bug:
                                    this.AddNewBugWorkItemAndSendReplace(request);
                                    break;
                                case MailAction.Replay:
                                    ReplayOnWIStateChange(request);
                                    break;
                                case MailAction.Ignore:
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "ERR", ex.Message));
                            this.TheMailObserver.SendLogErrorMail(string.Format("ERRORE MAIL TO BUG! Dettagli: {0}\n Stack: {1}", ex.Message, ex.StackTrace));
                        }
                    }

                    Monitor.PulseAll(this.MailQueue);
                }
            }
        }

        protected void ReplayOnWIStateChange(T request)
        {
            // qui ci va la logica di risposta
            var id = this.GetWIId(request);
            var wiStateByMail = GetWiStateByMail(request);
            var workItemData = this.WIManager.GetWiData(id);
            var replyData = this.GetMailData(workItemData.AttachmentPath, id);
            var response = string.Empty;
            //se non lo trovo nella mail uso quello di TFS
            workItemData.WIState = wiStateByMail ?? workItemData.WIState;
            switch (workItemData.WIState)
            {
                case "Done":
                    response = string.Format("Il Work Item ID: {0} è stato chiuso.\nQuello che segue è lo storico del Work Item in oggetto:\n{1}", id, workItemData.History);
                    replyData.Subject = string.Format("R: {0} - CHIUSO", replyData.Subject);
                    //this.TheMailObserver.Client.SendReplay(replyData.From, replyData.Subject, replyData.Attachment, response);
                    //l'attachment non serve più nella risposta (ma serve per il mittente)
                    this.TheMailObserver.Client.SendReplay(replyData.From, replyData.Subject, null, response);
                    break;
                case "Committed":
                    response = string.Format("Il Work Item ID: {0} è passato in lavorazione", id);
                    replyData.Subject = string.Format("R: {0} - IN LAVORAZIONE", replyData.Subject);
                    //in lavorazione non vogliono l'attachment
                    this.TheMailObserver.Client.SendReplay(replyData.From, replyData.Subject, null, response);
                    break;
                default:
                    throw new Exception(string.Format("Lo stato {0} del WI {1} non è gestito.", id, workItemData.WIState));
            }
        }

        private string GetWiStateByMail(T request)
        {
            var bodyRows = request.Body.Split('\n');
            for (int i = 0; i < bodyRows.Length; i++)
            {
                var m = Regex.Match(bodyRows[i], @"<td\sclass=""PropName"">State:", RegexOptions.IgnoreCase);
                if (m.Success && i + 1 < bodyRows.Length)
                {
                    m = Regex.Match(bodyRows[i + 1], @"<td\sclass=""PropValue"">(?<data>.+)</td>");
                    if (m.Success)
                    {
                        return m.Groups["data"].Value;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// finalizzazione del servizio
        /// </summary>
        public virtual void StopApplication()
        {
            try
            {
                for (int i = 0; i < this.Workers.Length; i++)
                {
                    var t = this.Workers[i];
                    if (t != null)
                    {
                        t.Abort();
                    }
                }

                for (int i = 0; i < this.MailClients.Length; i++)
                {
                    var t = this.MailClients[i];
                    if (t != null)
                    {
                        t.Abort();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(new LogInfo(System.Reflection.MethodBase.GetCurrentMethod(), "ERR", string.Format("Errore all'avvio del servizio. Dettagli: {0}", ex.Message)));
                throw ex;
            }
            finally
            {
                Logger.StopLoggingAndWait(1000);
            }
        }

        /// <summary>
        /// Elabora un messaggio aggiungendo un nuovo Work Item
        /// </summary>
        /// <param name="data">
        /// il messaggio da elaborare
        /// </param>
        /// <returns>
        /// la stringa da inoltrare al mittente
        /// </returns>
        protected virtual void AddNewBugWorkItemAndSendReplace(MessageRequest data)
        {
            var response = string.Empty;
            try
            {
                // data.OriginalMessage.CcRecipients.Clear();
                var id = this.WIManager.AddNewWIFromMailBug(data);
                response = string.Format("La tua richiesta è stata elaborata.\nL'ID del Work Item creato è {0}", id);
                this.TheMailObserver.Client.SendResponse(data.GetOriginalMessage(), response);

            }
            catch (Exception ex)
            {
                var errorMessage = string.Format("Si è verificato in problema durante la creazione del WorkItem. Dettagli: {0}", ex.Message);
                Logger.Error(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", errorMessage));
                response = string.Format("{0}\n{1}\n{2} {3}",
                    "C'è stato un problema durante l'elaborazione della richiesta",
                    "Verificare che l'allegato non superi le dimensioni massime (4,194,304 Byte)",
                    "Se il problema persiste contattare",
                    this.TheMailObserver.Client.AddressToMailOnError);
                this.TheMailObserver.Client.SendResponse(data.GetOriginalMessage(), response);
                throw ex;
            }
        }

        /// <summary>
        /// inizializzazione del sistema di log
        /// </summary>
        protected virtual void InitLogger()
        {
            if (!Logger.Initialized)
            {
                string config = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "loginfo.config");

                // using (var wrt = new StreamWriter("c:\\Cancellami.messaggio.txt"))
                // {
                //    wrt.WriteLine("Servizio email to bug");
                //    wrt.WriteLine(string.Format("directory corrente: {0}", Directory.GetCurrentDirectory()));
                //    wrt.WriteLine(string.Format("directory appdomain: {0}", AppDomain.CurrentDomain.BaseDirectory));
                // }
                Logger.Init(config);
            }
        }

        /// <summary>
        /// evento di start
        /// </summary>
        /// <param name="args">
        /// parametri (non usati)
        /// </param>
        protected override void OnStart(string[] args)
        {
            this.InitApplication();
            base.OnStart(args);
        }

        /// <summary>
        /// evento di stop
        /// </summary>
        protected override void OnStop()
        {
            this.StopApplication();
            base.OnStop();
        }

        /// <summary>
        /// Analizza l'email e decide le azioni da intraprendere
        /// </summary>
        /// <param name="request">
        /// il messaggio da analizzare
        /// </param>
        /// <returns>
        /// l'azione da intraprendere
        /// </returns>
        private MailAction CheckElaborateThisMail(MessageRequest request)
        {
            foreach (var str in this.automaticResponseString)
            {
                if (request.Subject.StartsWith(str))
                {
                    return MailAction.Ignore;
                }
            }

            foreach (var str in this.replyAddresses)
            {
                if (request.From.ToLower().CompareTo(str) == 0)
                {
                    return MailAction.Replay;
                }
            }

            return MailAction.Bug;
        }

        /// <summary>
        /// Ottiene il mittente della mail
        /// </summary>
        /// <param name="rows">
        /// le righe del file da analizzare
        /// </param>
        /// <param name="wiid">
        /// l'id del work item per la tracciatura dell'errore
        /// </param>
        /// <returns>
        /// il mittente della mail
        /// </returns>
        private string GetFromAddress(string[] rows, string wiid)
        {
            foreach (var r in rows)
            {
                var m = Regex.Match(r, this.regexMittenteMail, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    return m.Groups["data"].Value;
                }
            }

            throw new Exception(string.Format("Impossibile estrarre il Mittente dall'attachment {0} relativo al work item {1}", string.Join("\r\n", rows), wiid));
        }

        /// <summary>
        /// Analizza l'email allegata al WI
        /// </summary>
        /// <param name="att">
        /// il path del file temporaneo
        /// </param>
        /// <param name="wiid">
        /// l'id del Work Item per la tracciatura degli errori
        /// </param>
        /// <returns>
        /// I dati della mail allegata al work item
        /// </returns>
        private MailReplyData GetMailData(string att, string wiid)
        {
            try
            {
                var strs = string.Empty;
                var result = new MailReplyData();
                using (var rdr = new StreamReader(att))
                {
                    strs = rdr.ReadToEnd();
                    var rows = strs.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    rows = this.RemoveCodedRows(rows);

                    result.From = this.GetFromAddress(rows, wiid);
                    result.Subject = this.GetMailSubject(rows, wiid);
                }

                using (var f = File.OpenRead(att))
                {
                    result.Attachment = new byte[f.Length];
                    f.Read(result.Attachment, 0, result.Attachment.Length);
                }

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // faccio pulizia
                TempFileManager.ClearFile(att);
            }
        }

        /// <summary>
        /// Ottiene l'oggetto della mail
        /// </summary>
        /// <param name="rows">
        /// le righe del file da analizzare
        /// </param>
        /// <param name="wiid">
        /// l'id del work item per la tracciatura dell'errore
        /// </param>
        /// <returns>
        /// l'oggetto della mail
        /// </returns>
        internal string GetMailSubject(string[] rows, string wiid)
        {
            foreach (var r in rows)
            {
                var m = Regex.Match(r, this.regexOggettoMail, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    return m.Groups["data"].Value;
                }
            }

            throw new Exception(string.Format("Impossibile estrarre l'oggetto dall'attachment:\n{0}\n\nrelativo al work item: {1}", string.Join("\r\n", rows), wiid));
        }

        internal string[] RemoveCodedRows(string[] rows)
        {
            var appo = rows.ToList();
            for (int i = 0; i < appo.Count; i++)
            {
                var m = Regex.Match(appo[i], base64Pattern, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    appo[i - 1] = appo[i - 1] + DecodeFromBase64(m.Groups["data"].Value);
                    appo.RemoveAt(i);
                    i--;
                    continue;
                }
                else
                {
                    m = Regex.Match(appo[i], quotedPattern, RegexOptions.IgnoreCase);
                    if (m.Success)
                    {
                        appo[i - 1] = appo[i - 1] + m.Groups["data"].Value;
                        appo.RemoveAt(i);
                        i--;
                        continue;
                    }
                }
            }
            return appo.ToArray();
        }

        internal string DecodeFromBase64(string encoded)
        {
            var bytes = Convert.FromBase64String(encoded);
            return Encoding.UTF8.GetString(bytes);
        }


        /// <summary>
        /// Ricerca una stringa particolare all'interno della mail di ALERT generata da
        /// TFS e ne estrae l'ID del WorkItem di interesse
        /// </summary>
        /// <param name="data">
        /// il messaggio da analizzare
        /// </param>
        /// <returns>
        /// la stringa che contiene l'ID del WorkItem
        /// </returns>
        private string GetWIId(MessageRequest data)
        {
            var rows = data.Subject.Split(new string[] { "\n" }, StringSplitOptions.None);
            foreach (var r in rows)
            {
                var m = Regex.Match(r, this.regexPatternToWIId, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    return m.Groups["data"].Value;
                }
            }

            throw new Exception("Impossibile identificare l'ID del WorkItem");
        }

        /// <summary>
        /// Classe privata serve per l'analisi del'EML
        /// </summary>
        private class MailReplyData
        {
            /// <summary>
            /// il file da allegare
            /// </summary>
            public byte[] Attachment
            {
                get;
                set;
            }

            /// <summary>
            /// L'indirizzo del mittente
            /// </summary>
            public string From
            {
                get;
                set;
            }

            /// <summary>
            /// L'oggetto della mail
            /// </summary>
            public string Subject
            {
                get;
                set;
            }
        }
    }
}