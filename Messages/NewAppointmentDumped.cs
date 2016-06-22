namespace Messages {

    public class NewAppointmentDumped : IEvent
    {
        public string Id { get; set; }
        public string MimeContent { get; set; }
        public dynamic Appointment { get; set; }
    }
}
