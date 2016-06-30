using System;
using System.IO;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Exchange.WebServices.Data;
using EWSAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
using EWSItemId = Microsoft.Exchange.WebServices.Data.ItemId;
using EWSAppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
using EasyNetQ;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SimpleConfig;

using Messages;
using CsvTargetAccountsProvider;

namespace EchangeExporterProto
{
    class Program
    {
        private static ExporterConfiguration config;
        private static MailboxAccountsProvider accountsProvider = new MailboxAccountsProvider(',');
        private static readonly string ACCOUNTSFILE = "targets.csv";
        private static ICollection<EWSAppointmentType> singleAndRecurringMasterAppointmentTypes = new List<EWSAppointmentType> { EWSAppointmentType.RecurringMaster, EWSAppointmentType.Single };

        private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service", "MimeContent"),
            Error = (serializer, err) => err.ErrorContext.Handled = true,
        };

        static void Main(string[] args)
        {
            config = Configuration.Load<ExporterConfiguration>();

            if (String.IsNullOrWhiteSpace(config.MessageQueue.ConnectionString) && String.IsNullOrWhiteSpace(config.MessageQueue.Host))
            {
                Error("Could not find either a connection string or an host for MQ!");
                return;
            }
            
            var queueConf = CompleteQueueConfigWithDefaults();

            if (String.IsNullOrWhiteSpace(config.Credentials.Domain)
                || String.IsNullOrWhiteSpace(config.Credentials.Login)
                || String.IsNullOrWhiteSpace(config.Credentials.Password))
            {
                Error("Provided credentials are incomplete!");
                return;
            }

            if (String.IsNullOrWhiteSpace(queueConf.ConnectionString))
                queueConf.ConnectionString = String.Format("host={0};virtualHost={1};username={2};password={3}",
                    queueConf.Host, queueConf.VirtualHost, queueConf.Username, queueConf.Password);            

            ExchangeService service = ConnectToExchange(config.ExchangeServer, config.Credentials);

            // Get all mailbox accounts
            var accountsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ACCOUNTSFILE);
            var mailboxes = accountsProvider.GetFromCsvFile(accountsFilePath);

            using (var bus = RabbitHutch.CreateBus(queueConf.ConnectionString , serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))
            {
                foreach (var mailbox in mailboxes)
                {
                    Console.WriteLine("Dumping calendar items for account: {0} ...", mailbox.PrimarySmtpAddress);
                    ImpersonateQueries(service, mailbox.PrimarySmtpAddress);

                    var foundEvents = FindAllMeetings(service, mailbox.PrimarySmtpAddress);

                    foreach (var ev in foundEvents)
                    {
                        Console.WriteLine("Extracted with event #{0}. About to publish to {1}...", ev.Id, config.MessageQueue.Host);
                        bus.Publish(ev);
                    }
                }
                

            }
            Console.ReadLine();
        }

        private static void ImpersonateQueries(ExchangeService service, string primaryAddress)
        {
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, primaryAddress);
        }

        private static MessageQueue CompleteQueueConfigWithDefaults()
        {
            // Set default value when missing
            return new MessageQueue {
                Host = config.MessageQueue.Host,
                Username = config.MessageQueue.Username ?? "guest",
                Password = config.MessageQueue.Password ?? "guest",
                VirtualHost = config.MessageQueue.VirtualHost ?? "/",
                Port = config.MessageQueue.Port != 0 ? config.MessageQueue.Port : 5672,
                ConnectionString = config.MessageQueue.ConnectionString
            };
        }

        private static string DumpAvailablePropsToJson(EWSAppointment ev)
        {
            return JsonConvert.SerializeObject(ev, Formatting.Indented, serializerSettings);
        }

        private static IEnumerable<NewAppointmentDumped> FindAllMeetings(ExchangeService service, String primaryEmailAddress)
        {
            PropertySet includeMostProps = BuildAppointmentPropertySet();
            // TODO: fix FindItems paged results
            // Beware: there is a server-based limit in paged size length (<1000 items)
            // may have to iterate over paged-results eventually, using a loop and FindItemResults<Item>.MoreAvailable prop ?
            var appIdsView = new ItemView(int.MaxValue) {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType)
            };
            // var appIdsView = new CalendarView(DateTime.UtcNow.AddYears(-1), DateTime.UtcNow);

            IQueryable<EWSAppointment> mailboxAppointments = GetAllCalendars(service)
                .Where(cal => cal.DisplayName == "SubCalendar1" || cal.DisplayName == "SecondRootCalendar")
                // .Take(1) // TODO: DEBUG REMOVE ME
                .SelectMany(calendar => service.FindItems(calendar.Id, appIdsView))
                .Cast<EWSAppointment>().AsQueryable();

            //var testApp = mailboxAppointments.ToList();
            //var testReccMasters = testApp.Where(app => app.AppointmentType == AppointmentType.RecurringMaster).ToList();
            //var testOccApp = testApp.Where(app => app.AppointmentType == AppointmentType.Exception).ToList();
            //var testOcc2App = testApp.Where(app => app.AppointmentType == AppointmentType.Occurrence).ToList();

            var singleAndReccurringMasterAppointments = mailboxAppointments.Where(app => singleAndRecurringMasterAppointmentTypes.Contains(app.AppointmentType));

            var singleAndReccurringMasterAppointmentsWithContext = singleAndReccurringMasterAppointments
                .Select(app => EWSAppointment.Bind(service, app.Id, includeMostProps))
                .Select(app => new {
                    Mailbox = primaryEmailAddress,
                    Folder = app.ParentFolderId,
                    Appointment = app
                });

            var messagesForExportingSingleAndReccurenceAppointments = singleAndReccurringMasterAppointmentsWithContext
                .Select(appCtx => new NewAppointmentDumped {
                    Mailbox = appCtx.Mailbox,
                    FolderId = appCtx.Folder.UniqueId,
                    Id = appCtx.Appointment.Id.ToString(),
                    Appointment = AddMissingModifiedOccurencesAttendees(service, appCtx.Appointment),
                    SourceAsJson = JsonConvert.SerializeObject(appCtx.Appointment, Formatting.Indented, serializerSettings),
                    MimeContent = Encoding.GetEncoding(appCtx.Appointment.MimeContent.CharacterSet).GetString(appCtx.Appointment.MimeContent.Content)
                });

            return messagesForExportingSingleAndReccurenceAppointments;
        }

        private static SearchFilter SkipAppointmentsOfType(EWSAppointmentType typeToSkip)
        {
            // filters on AppointmentSchema.AppointmentType throw at runtime "The property can not be used with this type of restriction."
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(AppointmentSchema.AppointmentType, typeToSkip));
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection);
            return searchFilter;
        }

        private static IEnumerable<CalendarFolder> GetAllCalendars(ExchangeService service)
        {
            int folderViewSize = int.MaxValue; // Kids, dont do this at home !
            FolderView view = new FolderView(folderViewSize);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.FolderClass);
            view.Traversal = FolderTraversal.Deep;
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);

            return findFolderResults.OfType<CalendarFolder>();
        }

        private static PropertySet BuildAppointmentPropertySet()
        {
            return new PropertySet(
                            BasePropertySet.FirstClassProperties,
                            AppointmentSchema.AppointmentType,
                            AppointmentSchema.Body,
                            AppointmentSchema.RequiredAttendees, 
                            AppointmentSchema.OptionalAttendees,
                            AppointmentSchema.Categories,
                            AppointmentSchema.Culture,
                            AppointmentSchema.DateTimeCreated,
                            AppointmentSchema.DateTimeReceived,
                            AppointmentSchema.DateTimeSent,
                            AppointmentSchema.DisplayTo,
                            AppointmentSchema.DisplayCc,
                            AppointmentSchema.Duration,
                            AppointmentSchema.End,
                            AppointmentSchema.Start,
                            AppointmentSchema.StartTimeZone,
                            AppointmentSchema.Subject,
                // AppointmentSchema.TextBody, // EWS 2013 only
                            AppointmentSchema.TimeZone,
                            AppointmentSchema.MimeContent,
                            AppointmentSchema.ModifiedOccurrences,
                            AppointmentSchema.DeletedOccurrences,
                            AppointmentSchema.IsRecurring, 
                            AppointmentSchema.AppointmentType,
                            AppointmentSchema.When
                        );
        }

        private static Messages.Appointment Convert(EWSAppointment app)
        {
            // TODO: there should be a genuine mapping eventually here, as models may diverge !
            var serializedPayload = JsonConvert.SerializeObject(app, Formatting.Indented, serializerSettings);
            return JsonConvert.DeserializeObject<Messages.Appointment>(serializedPayload);
        }

        private static Messages.Appointment AddMissingModifiedOccurencesAttendees(ExchangeService service, EWSAppointment appointment)
        {
            if (appointment.AppointmentType != EWSAppointmentType.RecurringMaster)
                return Convert(appointment);

            var reccurenceExceptionIds = appointment.ModifiedOccurrences.Select(occ => occ.ItemId);
            var reccurenceExceptionsAttendees = reccurenceExceptionIds
                .Select(id => EWSAppointment.Bind(service, id, new PropertySet( 
                    BasePropertySet.IdOnly, AppointmentSchema.RequiredAttendees, AppointmentSchema.OptionalAttendees)))
                .ToDictionary( k => ConvertIdFrom(k.Id), v => v );

            var serializedPayload = JsonConvert.SerializeObject(appointment, Formatting.Indented, serializerSettings);
            var dtoAppointment = JsonConvert.DeserializeObject<Messages.Appointment>(serializedPayload);

            foreach (var occurence in dtoAppointment.ModifiedOccurrences)
            {
                if (reccurenceExceptionsAttendees.ContainsKey(occurence.ItemId))
                {
                    var exceptionAppointmentInfo = reccurenceExceptionsAttendees[occurence.ItemId];
                    // NOTE: ignoring optional attendees for now
                    occurence.Attendees = exceptionAppointmentInfo.RequiredAttendees
                        .Select(MapRequiredAttendees)
                        .ToList<Messages.Attendee>();
                }
            }
            return dtoAppointment;
        }

        private static RequiredAttendee MapRequiredAttendees(Microsoft.Exchange.WebServices.Data.Attendee att)
        {
            return new Messages.RequiredAttendee {
                Address = att.Address,
                Name = att.Name,
                MailboxType = (int?)att.MailboxType,
                ResponseType = (int?)att.ResponseType,
                RoutingType = att.RoutingType,
            };
        }

        private static Messages.ItemId ConvertIdFrom(EWSItemId id)
        {
            return new Messages.ItemId {
                uniqueId = id.UniqueId,
                changeKey = id.ChangeKey
            };
        }

        private static Messages.Appointment AddMissingAttendees(EWSAppointment appointment, Dictionary<Messages.ItemId, EWSAppointment> reccurenceExceptionsIndex)
        {
            var serializedPayload = JsonConvert.SerializeObject(appointment, Formatting.Indented, serializerSettings);
            var dtoAppointment = JsonConvert.DeserializeObject<Messages.Appointment>(serializedPayload);

            if (appointment.AppointmentType == EWSAppointmentType.RecurringMaster)
            {
                foreach (var occurence in dtoAppointment.ModifiedOccurrences)
                {
                    if (reccurenceExceptionsIndex.ContainsKey(occurence.ItemId))
                    {
                        var exceptionAppointmentInfo = reccurenceExceptionsIndex[occurence.ItemId];
                        // NOTE: ignoring optional attendees for now
                        occurence.Attendees = exceptionAppointmentInfo.RequiredAttendees
                            .Select(att => new Messages.RequiredAttendee {
                                Address = att.Address,
                                Name = att.Name,
                                MailboxType = (int?)att.MailboxType,
                                ResponseType = (int?)att.ResponseType,
                                RoutingType = att.RoutingType,
                            })
                            .ToList<Messages.Attendee>();
                    }
                }
            }
            return dtoAppointment;
        }

        private static ExchangeService ConnectToExchange(ExchangeServer exchange, Credentials credentials) {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            // Ignoring invalid exchange server provided certificate, on purpose, Yay !
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

            service.Credentials = new NetworkCredential(credentials.Login, credentials.Password, credentials.Domain);
            string exchangeEndpoint = String.Format(exchange.EndpointTemplate, exchange.Host);
            service.Url = new Uri(exchangeEndpoint);
            return service;
        }

        private static void Error(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
