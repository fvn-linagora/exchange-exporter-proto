﻿using System;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using EasyNetQ;

using Messages;
using Attendee = DDay.iCal.Attendee;
using ICalendarTransformersRegistry = System.Collections.Generic.IDictionary<Messages.AppointmentType, System.Func<DDay.iCal.IICalendar, Messages.Appointment, DDay.iCal.IICalendar>>;
using EventDateMapper = System.Func<DDay.iCal.IEvent, System.DateTimeOffset>;

namespace EchangeDumpedMessagesListener
{
    class Program
    {
        private static readonly string MQCONNETIONSTRING = ConfigurationManager.ConnectionStrings["spewsMQ"].ConnectionString;
        private static iCalendarSerializer serializer = new iCalendarSerializer();
        private static ICalendarTransformersRegistry mapOfICalendarTransformersByType = BuildCalendarTransformersRegistry();

        private const string EMAIL_VALIDATOR_PATTERN =
            @"[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?";

        private static readonly Regex emailValidationRegex = new Regex(EMAIL_VALIDATOR_PATTERN, RegexOptions.Compiled);

        static readonly Dictionary<Type, string> mapAttendeeTypeToICalendarRole = new Dictionary<Type, string>
        {
            [typeof(RequiredAttendee)] = "REQ-PARTICIPANT",
            [typeof(OptionalAttendee)] = "OPT-PARTICIPANT",
            [typeof(Resource)] = "NON-PARTICIPANT"
        };

        static Dictionary<Type, Func<Appointment, IEnumerable<InvitedAttendee>>> mapAttendeeTypeToCollectionPropertyProvider = new Dictionary<Type, Func<Appointment, IEnumerable<InvitedAttendee>>>
        {
            [typeof(RequiredAttendee)] = app => app.RequiredAttendees,
            [typeof(OptionalAttendee)] = app => app.OptionalAttendees,
            [typeof(Resource)] = app => app.Resources,
        };


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
                {AppointmentType.RecurringMaster, ReccuringOccurenceHandler},
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

            var fixCalendar = mapOfICalendarTransformersByType[app.Appointment.AppointmentType];
            var fixedCalendar = fixCalendar(iCal, app.Appointment);

            var eventMimeWithAttendees = serializer.SerializeToString(fixedCalendar);

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
            var missingAttendees = ConvertAttendeesOfType<RequiredAttendee>(appointment)
                .Union(ConvertAttendeesOfType<OptionalAttendee>(appointment))
                .Union(ConvertAttendeesOfType<Resource>(appointment));

            var eventWithAttendees = iCalEvent.Copy<Event>();
            eventWithAttendees.Attendees.Clear();
            eventWithAttendees.Attendees.AddRange(missingAttendees);

            if (eventWithAttendees.Organizer == null && appointment?.Organizer?.Address != null) {
                eventWithAttendees.Organizer = new DDay.iCal.Organizer {
                    CommonName = appointment.Organizer.Name,
                    Value = new Uri($"mailto:{ appointment.Organizer.Address}"),
                };
            }

            return eventWithAttendees;
        }

        private static IEnumerable<DDay.iCal.Attendee> ConvertAttendeesOfType<T>(Appointment appointment) where T : InvitedAttendee
        {
            var attendeeType = typeof(T);
            if (!mapAttendeeTypeToCollectionPropertyProvider.ContainsKey(attendeeType))
                yield break;

            var attendeesProvider = mapAttendeeTypeToCollectionPropertyProvider[attendeeType];
            var attendeesCollection = attendeesProvider(appointment);
            if (attendeesCollection == null)
                yield break;

            Func<InvitedAttendee, bool> isNotAnOrganizer = attendee => !IsAppointmentOrganizer(appointment, attendee);

            var iCalAttendees = attendeesCollection
                .Where(IsAttendeesAddressSet)
                .Where(isNotAnOrganizer)
                .Select(attendee => ConvertToICal<T>(appointment.IsResponseRequested, attendee));

            foreach (var attendee in iCalAttendees)
                yield return attendee;
        }

        private static DDay.iCal.Attendee ConvertToICal<T>(bool isResponseRequested, InvitedAttendee attendee) where T : InvitedAttendee
        {
            if (attendee == null || !IsAttendeesAddressSet(attendee))
                throw new ArgumentNullException(nameof(attendee));

            string contactAddress = HasAttendeeGotAnEmailSet(attendee) ? attendee.Address : attendee.Name;

            return new DDay.iCal.Attendee
            {
                Value = new Uri($"mailto:{contactAddress}"),
                CommonName = attendee.Name,
                ParticipationStatus = ConvertToParticipationStatus(attendee.ResponseType),
                Type = GuessUserType<T>(attendee),
                Role = ConvertToRole<T>(),
                RSVP = isResponseRequested,
            };
        }

        private static bool IsAppointmentOrganizer(Appointment appointment, InvitedAttendee attendee)
        {
            return string.Equals((appointment?.Organizer).Address, attendee.Address, StringComparison.InvariantCultureIgnoreCase);
        }

        private static string ConvertToRole<T>() where T : InvitedAttendee
        {
            return mapAttendeeTypeToICalendarRole.ContainsKey(typeof(T)) ? mapAttendeeTypeToICalendarRole[typeof(T)] : null;
        }

        private static string GuessUserType<T>(InvitedAttendee attendee) where T : InvitedAttendee
        {
            switch (typeof(T).Name)
            {
                case nameof(Resource):
                    return "RESOURCE";
                default:
                    switch(attendee.MailboxType)
                    {
                        case MailboxType.ContactGroup:
                        case MailboxType.PublicFolder:
                        case MailboxType.PublicGroup:
                            return "GROUP";
                        case MailboxType.Contact:
                        case MailboxType.Mailbox:
                            return "INDIVIDUAL";
                        default:
                            return "UNKNOWN";
                    }
            }
        }

        private static string ConvertToParticipationStatus(MeetingResponseType? responseType)
        {
            if (!responseType.HasValue)
                return "NEEDS-ACTION";

            switch(responseType.Value)
            {
                case MeetingResponseType.Accept:
                    return "ACCEPTED";
                case MeetingResponseType.Decline:
                    return "DECLINED";
                case MeetingResponseType.Tentative:
                    return "TENTATIVE";
                default:
                    return "NEEDS-ACTION";
            }
        }

        private static bool IsAttendeesAddressSet(Messages.Attendee attendee) =>
            HasAttendeeGotAnEmailSet(attendee) && !string.IsNullOrWhiteSpace(attendee.Address) || IsAttendeeNameAnAddress(attendee);
        private static bool IsAttendeeNameAnAddress(Messages.Attendee attendee) => !HasAttendeeGotAnEmailSet(attendee) && IsEmailAddress(attendee?.Name);
        private static bool HasAttendeeGotAnEmailSet(Messages.Attendee a) => a?.RoutingType?.Equals("SMTP", StringComparison.OrdinalIgnoreCase) ?? false;

        private static bool IsEmailAddress(string attendeeName) => !string.IsNullOrWhiteSpace(attendeeName) && emailValidationRegex.IsMatch(attendeeName);

        private static IICalendar ReccuringOccurenceHandler(IICalendar calendar, Messages.Appointment reccuringAppointment)
        {
            var calendarWithAttendees = AppendMissingAttendeesToReccuringOccurence(calendar, reccuringAppointment);
            var calendarWithAttendeesAndFixedRecurrenceIds = FixReccuringExceptionsReccurenceId(calendarWithAttendees, reccuringAppointment);

            return calendarWithAttendeesAndFixedRecurrenceIds;
        }

        private static IICalendar AppendMissingAttendeesToReccuringOccurence(IICalendar calendar, Messages.Appointment reccuringAppointment)
        {
            if (reccuringAppointment == null)
                return calendar;
            if (reccuringAppointment.ModifiedOccurrences == null)
                return calendar;

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
                    Event = vev,
                    Attendees = ConvertAttendeesToICal(att.Attendees, reccuringAppointment.IsResponseRequested)
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

        private static IEnumerable<DDay.iCal.Attendee> ConvertAttendeesToICal(IEnumerable<Messages.Attendee> attendees, bool isResponseRequested)
        {
            var requiredAttendees = attendees
                .OfType<RequiredAttendee>()
                .Select(attendee => ConvertToICal<RequiredAttendee>(isResponseRequested, attendee));
            var optionalAttendees = attendees
                .OfType<OptionalAttendee>()
                .Select(attendee => ConvertToICal<OptionalAttendee>(isResponseRequested, attendee));
            var resources = attendees
                .OfType<Resource>()
                .Select(attendee => ConvertToICal<Resource>(isResponseRequested, attendee));
            var result = requiredAttendees.Concat(optionalAttendees).Concat(resources);

            return result;
        }

        private static IICalendar FixReccuringExceptionsReccurenceId(IICalendar calendar, Appointment reccuringAppointment)
        {
            if (reccuringAppointment == null || reccuringAppointment.ModifiedOccurrences == null)
                return calendar;

            var updatedCalendar = calendar.Copy<IICalendar>();
            updatedCalendar.Events.Clear(); // All but events

            var calendarEvents = calendar.Events.ToList();
            var fixReccurrenceId = BuildEventRecurrenceIdFixer(reccuringAppointment, calendarEvents);

            calendarEvents
                .Select(fixReccurrenceId)
                .ToList()
                .ForEach(updatedCalendar.Events.Add);

            return updatedCalendar;
        }

        private static Func<DDay.iCal.IEvent, Event> BuildEventRecurrenceIdFixer(Appointment reccuringAppointment, List<DDay.iCal.IEvent> calendarEvents)
        {
            var mapOfModifiedOccurrences = BuildMapOfStartDateToReccurrenceException(calendarEvents, reccuringAppointment.ModifiedOccurrences);
            EventDateMapper originalStartDateMapper = ev => DateTimeOffset.Parse(mapOfModifiedOccurrences[ev.Start].OriginalStart);
            Func<DDay.iCal.IEvent, Event> fixReccurrenceId = ev => FixEventReccuringId(originalStartDateMapper, ev);
            return fixReccurrenceId;
        }

        private static IDictionary<IDateTime, ModifiedOccurrence> BuildMapOfStartDateToReccurrenceException(
            IEnumerable<DDay.iCal.IEvent> events, ICollection<ModifiedOccurrence> modifiedOccurrences)
        {
            var eventsJoin =
                from e in events
                join occ in modifiedOccurrences
                on e.Start.UTC equals DateTime.Parse(occ.Start).ToUniversalTime()
                where e.RecurrenceID != null
                select new { Start = e.Start, Exception = occ };

            return eventsJoin.ToDictionary(k => k.Start, v => v.Exception);
        }

        private static Event FixEventReccuringId(EventDateMapper originalStartDateProvider, DDay.iCal.IEvent ev)
        {
            if (ev.RecurrenceID == null && ev is Event)
                return ev as Event; // reccurence master
            var updatedEvent = ev.Copy<Event>();
            var eventOriginalStartUtcDate = originalStartDateProvider(ev).ToUniversalTime();
            var originalUtcDateICal = new iCalDateTime(eventOriginalStartUtcDate.DateTime) { IsUniversalTime = true };
            updatedEvent.RecurrenceID = originalUtcDateICal;

            return updatedEvent;
        }
    }
}
