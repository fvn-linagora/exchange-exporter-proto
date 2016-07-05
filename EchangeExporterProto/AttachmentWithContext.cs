using System;
using CsvTargetAccountsProvider;
using Microsoft.Exchange.WebServices.Data;

namespace EchangeExporterProto
{
    class AttachmentWithContext
    {
        public MailAccount Mailbox { get; set; }

        public CalendarFolder Calendar { get; set; }
    
        public Appointment Appointment { get; set; }

        public FileAttachment Attachment { get; set; }
    }
}
