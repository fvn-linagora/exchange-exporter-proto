namespace EchangeExporterProto
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Exchange.WebServices.Data;
    using AppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
    using Attendee = Microsoft.Exchange.WebServices.Data.Attendee;
    using EWSAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
    using EWSAppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;

    public class AppointmentWithParticipation
    {
        public AppointmentWithParticipation(EWSAppointment appointment)
        {
            if (appointment == null)
                throw new ArgumentNullException(nameof(appointment));
            Appointment = appointment;
            ExceptionsAttendees = new Dictionary<ItemId, ExceptionAttendees>();
        }

        public EWSAppointment Appointment { get; }
        public IDictionary<ItemId, ExceptionAttendees> ExceptionsAttendees { get; set; }
    }

    public class ExceptionAttendees
    {
        public AttendeeCollection Optional { get; set; }
        public AttendeeCollection Required { get; set; }
        public AttendeeCollection Resources { get; set; }
    }
}
