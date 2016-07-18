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

    public class NewEventAttachment : IEvent
    {
        public Guid Id { get; set; }
        public DateTimeOffset CreationDate { get; set; }
        public DateTimeOffset LastModified { get; set; }

        public string PrimaryEmailAddress { get; set; }
        public string CalendarId { get; set; }
        public string AppointmentId { get; set; }
        public byte[] Content { get; set; }
    }

    public class NewAddressBook : IEvent
    {
        public Guid Id { get; set; }
        public DateTimeOffset CreationDate { get; set; }
        public string PrimaryEmailAddress { get; set; }
        public string AddressBookId { get; set; }
        public string DisplayName { get; set; }
    }
}
