namespace Messages {

    public class NewAppointmentDumped : IEvent
    {
        public string Id { get; set; }
        public string Owner { get; set; }
        public string MimeContent { get; set; }
    }
}
