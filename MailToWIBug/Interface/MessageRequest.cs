using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToWIBug.Interface
{
    public interface MessageRequest
    {
        EmailMessage GetOriginalMessage();

        string From { get; set; }

        string To { get; set; }

        string Subject { get; set; }

        string Body { get; set; }

        byte[] FileFormat { get; set; }

        EmailMessage OriginalMessage { get; set; }


    }
}
