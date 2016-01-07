using MailToWIBug.Interface;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToWIBug.Code
{
    public class MailRequest : MessageRequest
    {

        public MailRequest(EmailMessage msg)
        {
            this.From = msg.From.Address;
            this.To = msg.ToRecipients.First().Address;
            this.Subject = msg.Subject;
            this.Body = msg.Body;
            this.OriginalMessage = msg;

            msg.Load(new PropertySet(ItemSchema.MimeContent));
            this.FileFormat = msg.MimeContent.Content;

        }

        public EmailMessage GetOriginalMessage()
        {
            return OriginalMessage;
        }



        public string From { get; set; }

        public string To { get; set; }

        public string Subject { get; set; }

        public string Body { get; set; }

        public EmailMessage OriginalMessage { get; set; }

        public byte[] FileFormat { get; set; }
    }
}
