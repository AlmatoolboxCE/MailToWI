using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailToWIBug;
using MailToWIBug.Code;

namespace MailToWIBug.DebugLauncher
{
    class Program
    {


        static void Main(string[] args)
        {
            MailToWIBugService<MailRequest> service = null;
            try
            {
                Console.WriteLine("MAIL to Work Item di tipo BUG");
                Console.WriteLine("Avvio del servizio in corso");
                service = new MailToWIBugService<MailRequest>();
                service.InitApplication();
                Console.WriteLine("Il servizio è partito");
                Console.WriteLine("Premere INVIO per uscire");
                Console.ReadLine();
            }
            catch (Exception ex)
            {

                Console.WriteLine(string.Format("Si è verificato un errore. Dettagli: {0}", ex.Message));
                Console.WriteLine("Premere INVIO per uscire");
                Console.ReadLine();
            }
            finally
            {
                try
                {                    
                    service.Dispose();
                }
                catch { }
                Console.WriteLine("Programma terminato");
            }

        }
    }
}
