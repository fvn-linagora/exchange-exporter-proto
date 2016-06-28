using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;

using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using EasyNetQ;

using Messages;

namespace EchangeDumpedMessagesListener
{
    class Program
    {
        private static readonly string MQCONNETIONSTRING = ConfigurationManager.ConnectionStrings["spewsMQ"].ConnectionString;
        private static iCalendarSerializer serializer = new iCalendarSerializer();
        
        static void Main(string[] args)
        {
            using (var bus = RabbitHutch.CreateBus(MQCONNETIONSTRING))
            {
                bus.Subscribe<NewAppointmentDumped>("test", CurryForwarderHandlerFrom(HandleNewAppointment, bus));

                Console.WriteLine("Listening for messages. Hit <return> to quit.");
                Console.ReadLine();
            }
        }


        static Action<NewAppointmentDumped> CurryForwarderHandlerFrom(Action<NewAppointmentDumped, Action<NewMimeEventExported>> forwarderHandler, IBus serviceBus) 
        {
            return incoming => forwarderHandler(incoming, serviceBus.Publish);
        }

        static void HandleNewAppointment(NewAppointmentDumped app, Action<NewMimeEventExported> newMessageConsumer)
        {
            Console.ForegroundColor = ConsoleColor.Red;

            var iCal = iCalendar.LoadFromStream(new StringReader(app.MimeContent)).Single(); // Should throw if Mime has multiple calendars/events

            if (app.Appointment.requiredAttendees.Count > 0 /* && app.Appointment.optionalAttendees > 0 */) 
            {
                var eventWithAttendees = iCal.Events.Single().Copy<Event>(); // Should throw if Mime has multiple calendars/events

                var missingAttendeesFromReceivedMime = app.Appointment.requiredAttendees
                    // .Union(app.Appointment.optionalAttendees)
                    .Where(IsAttendeesAddressSet)
                    .Select(a => new {
                        Uri = new Uri("mailto:" + a.Address),
                        DisplayName = a.Name,
                    })
                    .Select(a => new DDay.iCal.Attendee(a.Uri) {
                        CommonName = a.DisplayName
                    })
                    .ToList();

                eventWithAttendees.Attendees.Clear();
                eventWithAttendees.Attendees.AddRange(missingAttendeesFromReceivedMime);

                iCal.Events.Clear();
                iCal.Events.Add(eventWithAttendees);

                var eventMimeWithAttendees = serializer.SerializeToString(iCal);

                var message = new NewMimeEventExported {
                    Id = Guid.NewGuid(),
                    CreationDate = DateTimeOffset.UtcNow,
                    PrimaryAddress = app.Mailbox,
                    CalendarId = app.FolderId,
                    AppointmentId = app.Id,
                    MimeContent = eventMimeWithAttendees
                };

                newMessageConsumer(message);
            }

            Console.WriteLine("Got message: {0}", app.Appointment.subject);
            Console.ResetColor();
        }

        private static bool IsAttendeesAddressSet(Messages.Attendee a)
        {
            return a.RoutingType.Equals("SMTP", StringComparison.OrdinalIgnoreCase) && !String.IsNullOrWhiteSpace(a.Address);
        }
    }
}
