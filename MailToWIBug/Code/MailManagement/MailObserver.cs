using AlmaLogger;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailToWIBug.Code
{
    public class MailObserver<T>
    {

        private const int MAX_TENTATIVI = 10;

        public MailObserver(Queue<T> queue)
        {
            this.Queue = queue;
            this.Client = new MailClient<T>();

            int seconds;
            if (!int.TryParse(ConfigurationManager.AppSettings["Mail.Polling.Interval"], out seconds))
            {
                ConfigurationErrorsException e = new ConfigurationErrorsException(string.Format("Mail.Polling.Interval = {0} non e' un numerico", ConfigurationManager.AppSettings["Mail.Polling.Interval"]));
                Logger.Fatal(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", string.Format("Errore nel caricamento dell'intervallo di polling mail. Dettagli: {0}", e.Message)));
                throw e;
            }
            this.RefreshRate = seconds * 60 * 1000;
        }

        public virtual void MailClientLoop()
        {
            int tentativi = MAX_TENTATIVI;
            while (true)
            {
                try
                {
                    ICollection<T> requests = this.Client.GetRequests();

                    lock (this.Queue)
                    {
                        foreach (T request in requests)
                        {
                            this.Queue.Enqueue(request);
                        }
                        Monitor.PulseAll(this.Queue);
                        Monitor.Wait(this.Queue);
                    }

                    if (this.RefreshRate != 0)
                    {
                        Thread.Sleep(this.RefreshRate);
                    }
                    //tutto ok rimetto a posto i tentativi
                    tentativi = MAX_TENTATIVI;

                    this.Client.CleanEmail();
                }
                catch (Exception ex)
                {
                    Logger.Error(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", string.Format("Si è verificato un errore in fase di ascolto delle mail in arrivo. Numero di tentativi rimasti: {1}. Dettagli: {0}", ex.Message, tentativi)));
                    if (tentativi > 0)
                    {
                        //se il server non risponde in 3 tentativi crash del servizio
                        tentativi--;
                        Thread.Sleep(this.RefreshRate);
                    }
                    else
                    {
                        Logger.Fatal(new LogInfo(MethodBase.GetCurrentMethod(), "ERR", string.Format("Si è verificato un errore in fase di ascolto delle mail in arrivo. Il Servizio sarà terminato. Dettagli: {0}", ex.Message)));
                        //se il server non risponde in 3 tentativi crash del servizio
                        throw ex;
                    }
                }
            }
        }


        public MailClient<T> Client { get; set; }

        public Queue<T> Queue { get; set; }

        public int RefreshRate { get; set; }

        public void SendLogErrorMail(string text)
        {
            this.Client.SendLogErrorMail(text);
        }
    }
}
