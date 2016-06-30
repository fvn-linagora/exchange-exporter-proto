namespace Messages {

    using System;

    public class NewAppointmentDumped : IEvent
    {
        public string Mailbox { get; set; }
        public string FolderId { get; set; }
        public string Id { get; set; }
        public string MimeContent { get; set; }
        public string SourceAsJson { get; set; }
        public Appointment Appointment { get; set; }
    }

    public class NewMimeEventExported : IEvent
    {
        public Guid Id { get; set; }
        public DateTimeOffset CreationDate { get; set; }
        public string PrimaryAddress { get; set; }
        public string CalendarId { get; set; }
        public string AppointmentId { get; set; }
        public string MimeContent { get; set; }
    }
}
