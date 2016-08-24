using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;

using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using EasyNetQ;

using Messages;
using ICalendarTransformersRegistry = System.Collections.Generic.IDictionary<Messages.AppointmentType, System.Func<DDay.iCal.IICalendar, Messages.Appointment, DDay.iCal.IICalendar>>;

namespace EchangeDumpedMessagesListener
{
    class Program
    {
        private static readonly string MQCONNETIONSTRING = ConfigurationManager.ConnectionStrings["spewsMQ"].ConnectionString;
        private static iCalendarSerializer serializer = new iCalendarSerializer();
        private static ICalendarTransformersRegistry mapOfICalendarTransformersByType = BuildCalendarTransformersRegistry();        

        static void Main(string[] args)
        {
            using (var bus = RabbitHutch.CreateBus(MQCONNETIONSTRING))
            {
                bus.Subscribe<NewAppointmentDumped>("test", CurryForwarderHandlerFrom(HandleNewAppointment, bus));

                Console.WriteLine("Listening for messages. Hit <return> to quit.");
                Console.ReadLine();
            }
        }


        private static ICalendarTransformersRegistry BuildCalendarTransformersRegistry()
        {
            return new Dictionary<AppointmentType, Func<IICalendar, Messages.Appointment, IICalendar>>() {
                {AppointmentType.Single, SingleEventAttendeesAppender},
                {AppointmentType.RecurringMaster, ReccuringOccurenceAttendeesAppender},
                {AppointmentType.Exception, (cal, _) => cal },
                {AppointmentType.Occurrence, (cal, _) => cal },
            };
        }


        static Action<NewAppointmentDumped> CurryForwarderHandlerFrom(Action<NewAppointmentDumped, Action<NewMimeEventExported>> forwarderHandler, IBus serviceBus) 
        {
            return incoming => forwarderHandler(incoming, serviceBus.Publish);
        }

        static void HandleNewAppointment(NewAppointmentDumped app, Action<NewMimeEventExported> newMessageConsumer)
        {
            Console.ForegroundColor = ConsoleColor.Red;

            var iCal = iCalendar.LoadFromStream(new StringReader(app.MimeContent)).Single(); // Should throw if Mime has multiple calendars/events

            var updateICalendar = mapOfICalendarTransformersByType[app.Appointment.AppointmentType]
                .Invoke(iCal, app.Appointment);

            var eventMimeWithAttendees = serializer.SerializeToString(updateICalendar);

            var message = new NewMimeEventExported {
                Id = Guid.NewGuid(),
                CreationDate = DateTimeOffset.UtcNow,
                PrimaryAddress = app.Mailbox,
                CalendarId = app.FolderId,
                AppointmentId = app.Id,
                MimeContent = eventMimeWithAttendees
            };

            newMessageConsumer(message);

            Console.WriteLine("Got message: {0}", app.Appointment.Subject);
            Console.ResetColor();
        }

        private static IICalendar SingleEventAttendeesAppender(IICalendar calendar, Messages.Appointment singleOccurenceAppointment)
        {
            if (singleOccurenceAppointment.RequiredAttendees.Count <= 0 && singleOccurenceAppointment.OptionalAttendees.Count <= 0)
                return calendar;

            var updatedEventWithAttendees = AddMissingAttendees(singleOccurenceAppointment, calendar.Events.Single().Copy<Event>());
            var updatedCalendar = calendar.Copy<IICalendar>();
            updatedCalendar.Events.Clear();
            updatedCalendar.Events.Add(updatedEventWithAttendees);
            return updatedCalendar;
        }

        private static DDay.iCal.IEvent AddMissingAttendees(Messages.Appointment appointment, DDay.iCal.IEvent iCalEvent)
        {
            var missingAttendees = appointment.RequiredAttendees
                .MapToICalAttendees()
                .ToList();

            var eventWithAttendees = iCalEvent.Copy<Event>();
            eventWithAttendees.Attendees.Clear();
            eventWithAttendees.Attendees.AddRange(missingAttendees);
            return eventWithAttendees;
        }

        private static IICalendar ReccuringOccurenceAttendeesAppender(IICalendar calendar, Messages.Appointment reccuringAppointment)
        {
            var updatedCalendar = calendar.Copy<IICalendar>();
            updatedCalendar.Events.Clear(); // All but events

            var exceptionsAttendeesOrderedByDate = reccuringAppointment.ModifiedOccurrences
                .Select(occ => new {
                    Start = DateTime.Parse(occ.Start).ToString("o"),
                    Attendees = occ.Attendees
                })
                .OrderBy(occ => occ.Start);

            var exceptionEventsWithAttendees = calendar.Events
                .Where(ev => ev.RecurrenceID != null)
                .OrderBy(ev => ev.Start.ToString("o"))
                .Zip(exceptionsAttendeesOrderedByDate, (vev, att) => new { 
                    Event = vev.Copy<Event>(),
                    Attendees = att.Attendees.MapToICalAttendees(),
                });

            var updatedExceptionEvents = exceptionEventsWithAttendees
                .Select(ewa => {
                    ewa.Event.Attendees.AddRange(ewa.Attendees);
                    return ewa.Event;
                });

            updatedExceptionEvents.ToList()
                .ForEach(updatedCalendar.Events.Add);

            var masterEvent = calendar.Events.Single(ev => ev.RecurrenceID == null).Copy<Event>();
            var updatedMasterEventWithAttendees = AddMissingAttendees(reccuringAppointment, masterEvent);
            updatedCalendar.Events.Add(updatedMasterEventWithAttendees);

            return updatedCalendar;
        }

        private static bool IsAttendeesAddressSet(Messages.Attendee a)
        {
            return a.RoutingType.Equals("SMTP", StringComparison.OrdinalIgnoreCase) && !String.IsNullOrWhiteSpace(a.Address);
        }
    }
}
