using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using EasyNetQ;
using CommandLine;
using Ical.Net;
using Ical.Net.DataTypes;
using Ical.Net.Interfaces;
using Ical.Net.Serialization.iCalendar.Serializers;
using IcalIEvent = Ical.Net.Interfaces.Components.IEvent;
using IcalEvent = Ical.Net.Event;
using IcalIAttendee = Ical.Net.Interfaces.DataTypes.IAttendee;
using IcalAttendee = Ical.Net.DataTypes.Attendee;

using Messages;
using ICalendarTransformersRegistry = System.Collections.Generic.IDictionary<
    Messages.AppointmentType,
    System.Func<Ical.Net.Interfaces.ICalendar, Messages.Appointment, Ical.Net.Interfaces.ICalendar>>;
using EventDateMapper = System.Func<Ical.Net.Interfaces.Components.IEvent, System.DateTimeOffset>;
using EchangeExporterProto;

namespace EchangeDumpedMessagesListener
{
    class Program
    {
        private const string EXPORTER_CONFIG_SECTION = "exporterConfiguration";
        private static readonly CalendarSerializer serializer = new CalendarSerializer();
        private static readonly ICalendarTransformersRegistry mapOfICalendarTransformersByType = BuildCalendarTransformersRegistry();
        private static readonly AutoResetEvent closingConsole = new AutoResetEvent(false);

        private const string EMAIL_VALIDATOR_PATTERN =
            @"[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?";

        private static readonly Regex emailValidationRegex = new Regex(EMAIL_VALIDATOR_PATTERN, RegexOptions.Compiled);

        static readonly Dictionary<Type, string> mapAttendeeTypeToICalendarRole = new Dictionary<Type, string>
        {
            [typeof(RequiredAttendee)] = "REQ-PARTICIPANT",
            [typeof(OptionalAttendee)] = "OPT-PARTICIPANT",
            [typeof(Resource)] = "NON-PARTICIPANT"
        };

        static readonly Dictionary<Type, Func<Appointment, IEnumerable<InvitedAttendee>>> mapAttendeeTypeToCollectionPropertyProvider= new Dictionary<Type, Func<Appointment, IEnumerable<InvitedAttendee>>>
        {
            [typeof(RequiredAttendee)] = app => app.RequiredAttendees,
            [typeof(OptionalAttendee)] = app => app.OptionalAttendees,
            [typeof(Resource)] = app => app.Resources,
        };


        static void Main(string[] args)
        {
            using (var bus = RabbitHutch.CreateBus(GetQueueConnectionString(args)))
            {
                bus.Subscribe("test", CurryForwarderHandlerFrom(HandleNewAppointment, bus));

                Console.WriteLine("Listening for messages...");
                Console.CancelKeyPress += OnExit;
                closingConsole.WaitOne();
            }
        }

        protected static void OnExit(object sender, ConsoleCancelEventArgs args)
        {
            Console.WriteLine("Exit");
            closingConsole.Set();
        }

        private static string GetQueueConnectionString(IEnumerable<string> args)
        {
            var result = Parser.Default.ParseArguments<ConfigurationOptions>(args);
            var arguments = result.MapResult(options => options, ArgumentErrorHandler);
            var config = new SimpleConfig.Configuration(configPath: arguments.ConfigPath)
                .LoadSection<ExporterConfiguration>(EXPORTER_CONFIG_SECTION);

            return BuildQueueConnectionString(config);
        }

        private static string BuildQueueConnectionString(ExporterConfiguration config)
        {
            if (string.IsNullOrWhiteSpace(config.MessageQueue.ConnectionString) && string.IsNullOrWhiteSpace(config.MessageQueue.Host))
                throw new Exception("Could not find either a connection string or an host for MQ!");

            var host = config.MessageQueue.Host;
            var username = config.MessageQueue.Username ?? "guest";
            var password = config.MessageQueue.Password ?? "guest";
            var virtualHost = config.MessageQueue.VirtualHost ?? "/";
            var port = config.MessageQueue.Port != 0 ? config.MessageQueue.Port : 5672;
            var connectionString = config.MessageQueue.ConnectionString;

            return !string.IsNullOrWhiteSpace(connectionString) ? connectionString
                : $"host={host}:{port};virtualHost={virtualHost};username={username};password={password}";
        }

        private static ConfigurationOptions ArgumentErrorHandler(IEnumerable<Error> errors)
        {
            Console.WriteLine($"Found issues with '{string.Join("\n", errors)}'");
            Environment.Exit(1);
            return default(ConfigurationOptions);
        }


        private static ICalendarTransformersRegistry BuildCalendarTransformersRegistry()
        {
            return new Dictionary<AppointmentType, Func<ICalendar, Appointment, ICalendar>> {
                {AppointmentType.Single, SingleEventAttendeesAppender},
                {AppointmentType.RecurringMaster, ReccuringOccurenceHandler},
                {AppointmentType.Exception, (cal, _) => cal },
                {AppointmentType.Occurrence, (cal, _) => cal },
            };
        }


        private static Action<NewAppointmentDumped> CurryForwarderHandlerFrom(Action<NewAppointmentDumped, Action<NewMimeEventExported>> forwarderHandler,
            IBus serviceBus) => incoming => forwarderHandler?.Invoke(incoming, serviceBus.Publish);

        static void HandleNewAppointment(NewAppointmentDumped app, Action<NewMimeEventExported> newMessageConsumer)
        {
            if (newMessageConsumer == null)
                throw new ArgumentNullException(nameof(newMessageConsumer));

            Console.ForegroundColor = ConsoleColor.Red;

            ICalendar iCal;
            using (var mimeStreamReader = new StringReader(app.MimeContent))
                iCal = Calendar.LoadFromStream<Calendar>(mimeStreamReader).SingleOrDefault();

            if (iCal == default(ICalendar)) {
                Console.WriteLine($"ERROR: Mailbox: '{app.Mailbox}', Event '{app.Appointment.Subject} organized by '{app.Appointment.Organizer}': found more than one calendar in MIME! Skipping ...");
                return;
            }
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

            Console.WriteLine($"Got message: {app.Appointment.Subject}");
            Console.ResetColor();
        }

        private static ICalendar SingleEventAttendeesAppender(ICalendar calendar, Appointment singleOccurenceAppointment)
        {
            if (singleOccurenceAppointment.RequiredAttendees.Count <= 0 && singleOccurenceAppointment.OptionalAttendees.Count <= 0)
                return calendar;

            var updatedEventWithAttendees = AddMissingAttendees(singleOccurenceAppointment, calendar.Events.Single().Copy<Event>());
            var updatedCalendar = calendar.Copy<ICalendar>();
            updatedCalendar.Events.Clear();
            updatedCalendar.Events.Add(updatedEventWithAttendees);
            return updatedCalendar;
        }

        private static IcalIEvent AddMissingAttendees(Appointment appointment, IcalIEvent iCalEvent)
        {
            var missingAttendeesButTheOrganizer = ConvertAttendeesOfType<RequiredAttendee>(appointment)
                .Union(ConvertAttendeesOfType<OptionalAttendee>(appointment))
                .Union(ConvertAttendeesOfType<Resource>(appointment))
                .Where(attendee => !IsEventOrganizer(appointment, attendee));

            var eventWithAttendees = iCalEvent.Copy<IcalEvent>();
            eventWithAttendees.Attendees.Clear();
            missingAttendeesButTheOrganizer.ToList().ForEach(eventWithAttendees.Attendees.Add);

            if (eventWithAttendees.Organizer == null && appointment?.Organizer?.Address != null) {
                eventWithAttendees.Organizer = new Ical.Net.DataTypes.Organizer {
                    CommonName = appointment.Organizer.Name,
                    Value = new Uri($"mailto:{ appointment.Organizer.Address}"),
                };
            }

            return eventWithAttendees;
        }

        private static bool IsEventOrganizer(Appointment appointment, IcalIAttendee attendee)
        {
            var attendeesAddress = $"{attendee.Value.UserInfo}@{attendee.Value.DnsSafeHost}";

            return string.Equals(appointment.Organizer.Address, attendeesAddress, StringComparison.InvariantCultureIgnoreCase);
        }

        private static IEnumerable<IcalIAttendee> ConvertAttendeesOfType<T>(Appointment appointment) where T : InvitedAttendee
        {
            var attendeeType = typeof(T);
            if (!mapAttendeeTypeToCollectionPropertyProvider.ContainsKey(attendeeType))
                yield break;

            var attendeesProvider = mapAttendeeTypeToCollectionPropertyProvider[attendeeType];
            var attendeesCollection = attendeesProvider(appointment);
            if (attendeesCollection == null)
                yield break;

            var iCalAttendees = attendeesCollection
                .Where(IsAttendeesAddressSet)
                .Select(attendee => ConvertToICal<T>(appointment.IsResponseRequested, attendee));

            foreach (var attendee in iCalAttendees)
                yield return attendee;
        }

        private static IcalAttendee ConvertToICal<T>(bool isResponseRequested, InvitedAttendee attendee) where T : InvitedAttendee
        {
            if (attendee == null || !IsAttendeesAddressSet(attendee))
                throw new ArgumentNullException(nameof(attendee));

            string contactAddress = HasAttendeeGotAnEmailSet(attendee) ? attendee.Address : attendee.Name;

            return new IcalAttendee
            {
                Value = new Uri($"mailto:{contactAddress}"),
                CommonName = attendee.Name,
                ParticipationStatus = ConvertToParticipationStatus(attendee.ResponseType),
                Type = GuessUserType<T>(attendee),
                Role = ConvertToRole<T>(),
                Rsvp = isResponseRequested,
            };
        }

        private static bool IsAppointmentOrganizer(Appointment appointment, Messages.Attendee attendee) =>
            string.Equals((appointment?.Organizer).Address, attendee.Address, StringComparison.InvariantCultureIgnoreCase);

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
                        case MailboxType.Unknown:
                        case MailboxType.OneOff:
                        case null:
                            return "UNKNOWN";
                        default:
                            return "UNKNOWN";
                    }
            }
        }

        public static string ConvertToParticipationStatus(MeetingResponseType? responseType)
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
                case MeetingResponseType.Unknown:
                case MeetingResponseType.Organizer:
                case MeetingResponseType.NoResponseReceived:
                    return "NEEDS-ACTION";
                default:
                    return "NEEDS-ACTION";
            }
        }

        private static bool IsAttendeesAddressSet(Messages.Attendee attendee) =>
            HasAttendeeGotAnEmailSet(attendee) && !string.IsNullOrWhiteSpace(attendee.Address) || IsAttendeeNameAnAddress(attendee);
        private static bool IsAttendeeNameAnAddress(Messages.Attendee attendee) => !HasAttendeeGotAnEmailSet(attendee) && IsEmailAddress(attendee?.Name);
        private static bool HasAttendeeGotAnEmailSet(Messages.Attendee a) => a?.RoutingType?.Equals("SMTP", StringComparison.OrdinalIgnoreCase) ?? false;

        private static bool IsEmailAddress(string attendeeName) => !string.IsNullOrWhiteSpace(attendeeName) && emailValidationRegex.IsMatch(attendeeName);

        private static ICalendar ReccuringOccurenceHandler(ICalendar calendar, Messages.Appointment reccuringAppointment)
        {
            // TODO: curry with reccuringAppointment and pipeline with calendars
            var calendarWithAttendees = AppendMissingAttendeesToReccuringOccurence(reccuringAppointment, calendar);
            var calendarWithAttendeesAndFixedRecurrenceIds = FixReccuringExceptionsReccurenceId(reccuringAppointment, calendarWithAttendees);
            var calendarWithAttendeesAndFixedRecurrenceIdsAndRepeatedOrganizers = RepeatOrganizerInOccurences(reccuringAppointment, calendarWithAttendeesAndFixedRecurrenceIds);

            return calendarWithAttendeesAndFixedRecurrenceIdsAndRepeatedOrganizers;
        }

        private static ICalendar AppendMissingAttendeesToReccuringOccurence(Appointment reccuringAppointment, ICalendar calendar)
        {
            if (reccuringAppointment == null)
                return calendar;
            if (reccuringAppointment.ModifiedOccurrences == null)
                return calendar;

            var updatedCalendar = calendar.Copy<ICalendar>();
            updatedCalendar.Events.Clear(); // All but events

            Func<Messages.Attendee, bool> isNotEventOrganizer = attendee => !IsAppointmentOrganizer(reccuringAppointment, attendee);

            var exceptionsAttendeesOrderedByDate = reccuringAppointment.ModifiedOccurrences
                .Select(occ => new {
                    Start = DateTime.Parse(occ.Start).ToString("o"),
                    Attendees = occ.Attendees.All.Where(isNotEventOrganizer)
                })
                .OrderBy(occ => occ.Start);

            var exceptionEventsWithAttendees = calendar.Events
                .Where(ev => ev.RecurrenceId != null)
                .OrderBy(ev => ev.Start.AsUtc.ToString("o"))
                .Zip(exceptionsAttendeesOrderedByDate, (vev, att) => new {
                    Event = vev,
                    Attendees = ConvertAttendeesToICal(att.Attendees, reccuringAppointment.IsResponseRequested)
                });

            var updatedExceptionEvents = exceptionEventsWithAttendees
                .Select(ewa => {
                    ewa.Attendees.ToList()
                        .ForEach(ewa.Event.Attendees.Add);
                    return ewa.Event;
                });

            updatedExceptionEvents.ToList()
                .ForEach(updatedCalendar.Events.Add);

            var masterEvent = calendar.Events.Single(ev => ev.RecurrenceId == null).Copy<IcalEvent>();
            var updatedMasterEventWithAttendees = AddMissingAttendees(reccuringAppointment, masterEvent);
            updatedCalendar.Events.Add(updatedMasterEventWithAttendees);

            return updatedCalendar;
        }

        private static IEnumerable<IcalAttendee> ConvertAttendeesToICal(IEnumerable<Messages.Attendee> attendees, bool isResponseRequested)
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

        private static ICalendar FixReccuringExceptionsReccurenceId(Appointment reccuringAppointment, ICalendar calendar)
        {
            if (reccuringAppointment?.ModifiedOccurrences == null)
                return calendar;

            var updatedCalendar = calendar.Copy<ICalendar>();
            updatedCalendar.Events.Clear(); // All but events

            var calendarEvents = calendar.Events.ToList();
            var fixReccurrenceId = BuildEventRecurrenceIdFixer(reccuringAppointment, calendarEvents);

            calendarEvents
                .Select(fixReccurrenceId)
                .ToList()
                .ForEach(updatedCalendar.Events.Add);

            return updatedCalendar;
        }

        private static Func<IcalIEvent, IcalEvent> BuildEventRecurrenceIdFixer(Appointment reccuringAppointment, IEnumerable<IcalIEvent> calendarEvents)
        {
            var mapOfModifiedOccurrences = BuildMapOfStartDateToReccurrenceException(calendarEvents, reccuringAppointment.ModifiedOccurrences);
            EventDateMapper originalStartDateMapper = ev => DateTimeOffset.Parse(mapOfModifiedOccurrences[ev.Start].OriginalStart);
            Func<IcalIEvent, IcalEvent> fixReccurrenceId = ev => FixEventReccuringId(originalStartDateMapper, ev);
            return fixReccurrenceId;
        }

        private static IDictionary<Ical.Net.Interfaces.DataTypes.IDateTime, ModifiedOccurrence> BuildMapOfStartDateToReccurrenceException(
            IEnumerable<IcalIEvent> events, ICollection<ModifiedOccurrence> modifiedOccurrences)
        {
            var eventsJoin =
                from e in events
                join occ in modifiedOccurrences
                    on e.Start.AsUtc equals DateTime.Parse(occ.Start).ToUniversalTime()
                where e.RecurrenceId != null
                select new { Start = e.Start, Exception = occ };

            return eventsJoin.ToDictionary(k => k.Start, v => v.Exception);
        }

        private static IcalEvent FixEventReccuringId(EventDateMapper originalStartDateProvider, IcalIEvent ev)
        {
            if (originalStartDateProvider == null)
                throw new ArgumentNullException(nameof(originalStartDateProvider));

            if (ev.RecurrenceId == null && ev is IcalEvent)
                return ev as IcalEvent; // reccurence master
            var updatedEvent = ev.Copy<Event>();
            var eventOriginalStartUtcDate = originalStartDateProvider(ev).ToUniversalTime();
            var originalUtcDateICal = new CalDateTime(eventOriginalStartUtcDate.DateTime) { IsUniversalTime = true };
            updatedEvent.RecurrenceId = originalUtcDateICal;

            return updatedEvent;
        }
        private static ICalendar RepeatOrganizerInOccurences(Appointment reccuringAppointment, ICalendar calendar)
        {
            if (reccuringAppointment?.ModifiedOccurrences == null)
                return calendar;
            if (string.IsNullOrWhiteSpace(reccuringAppointment.Organizer?.Address))
                return calendar;

            var updatedCalendar = calendar.Copy<ICalendar>();

            var foundOrganizer = calendar.Events.FirstOrDefault(ev => ev.Organizer != null).Organizer;

            foreach (var ev in updatedCalendar.Events)
                ev.Organizer = foundOrganizer;

            return updatedCalendar;
        }
    }
}
