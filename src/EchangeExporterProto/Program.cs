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
using SimpleConfig;
using CommandLine;

using Messages;
using CsvTargetAccountsProvider;

namespace EchangeExporterProto
{
    class Options
    {
        // [Option('t', "targets", Required = true,
        [Option('t', "targets",
          HelpText = "Input mailboxes to be processed.")]
        public string TargetsListFile { get; set; }

        [Option('c', "config",
           HelpText = "Configuration file path.")]
        public string ConfigPath { get; set; }

        [Option(
          HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }
    }

    class Program
    {
        private const string EXPORTER_CONFIG_SECTION = "exporterConfiguration";
        private const string DEFAULT_EXPORTER_CONFIG = "ExchangeExporter.config";
        private const string ENV_EXPORTER_CONFIG = "EXPORTER_CONFIG";
        private static ExporterConfiguration config;
        private static MailboxAccountsProvider accountsProvider = new MailboxAccountsProvider(',');
        private static readonly string ACCOUNTSFILE = "targets.csv";
        private static ICollection<EWSAppointmentType> singleAndRecurringMasterAppointmentTypes = new List<EWSAppointmentType> { EWSAppointmentType.RecurringMaster, EWSAppointmentType.Single };

        private static readonly ILog log = new ConsoleLogger();

        private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service", "MimeContent"),
            Error = (serializer, err) => err.ErrorContext.Handled = true,
        };

        static void Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<Options>(args);
            var arguments = result.MapResult( options => options, ArgumentErrorHandler);

            config = new Configuration(configPath: GetConfigPath(arguments)).LoadSection<ExporterConfiguration>(EXPORTER_CONFIG_SECTION);
            var mailboxes = GetTargetAccounts(arguments).ToList();

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

            ExportAndPublishAppointments(queueConf, service, mailboxes);
            ExportAndPublishAddressBooks(queueConf, service, mailboxes);
            var attachedMessages = ExportAppointmentsAttachedFiles(service, mailboxes).Select(MapToAttachmentMessage);
            PublishToBus(attachedMessages, queueConf);

            Console.ReadLine();
        }

        private static Options ArgumentErrorHandler(IEnumerable<Error> errors) {
            log.Error( String.Format("Found issues with '{0}'", String.Join("\n", errors) ));
            Environment.Exit(1);
            return default(Options);
        }

    private static string GetConfigPath(Options arguments)
        {
            string configPath =
                // First check args
                (String.IsNullOrWhiteSpace(arguments.ConfigPath) ? null : arguments.ConfigPath)
                // Then ENV variable
                ?? Environment.GetEnvironmentVariable(ENV_EXPORTER_CONFIG)
                // Then get default path
                ?? DEFAULT_EXPORTER_CONFIG;

            return Path.GetFullPath(configPath);
        }

        private static NewEventAttachment MapToAttachmentMessage(AttachmentWithContext attachment)
        {
            if (attachment.Attachment.Content == null)
                attachment.Attachment.Load();

            return new NewEventAttachment {
                Id = Guid.NewGuid(),
                CreationDate = DateTime.UtcNow,
                LastModified = attachment.Appointment.LastModifiedTime,
                PrimaryEmailAddress = attachment.Mailbox.PrimarySmtpAddress,
                CalendarId = attachment.Calendar.Id.UniqueId,
                AppointmentId = attachment.Appointment.Id.UniqueId,
                Content = attachment.Attachment.Content
            };
        }

        private static void PublishToBus<T>(IEnumerable<T> messages, MessageQueue queueConf) where T: class, new()
        {
            using (var bus = RabbitHutch.CreateBus(queueConf.ConnectionString ,
                serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))
            {
                foreach (var message in messages)
                {
                    bus.Publish(message);
                }
            }
        }

        private static IEnumerable<MailAccount> GetTargetAccounts(Options arguments)
        {
            var accountsFilePath = FindTargetAccountsFile(arguments) ??
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ACCOUNTSFILE);
            return accountsProvider.GetFromCsvFile(accountsFilePath);
        }

        private static string FindTargetAccountsFile(Options arguments)
        {
            if (String.IsNullOrWhiteSpace(arguments.TargetsListFile))
                return null;
            if (Path.GetFullPath(arguments.TargetsListFile) == null)
                return null;
            if (!File.Exists(arguments.TargetsListFile))
                return null;
            return arguments.TargetsListFile;
        }

        private static IEnumerable<AttachmentWithContext> ExportAppointmentsAttachedFiles(ExchangeService service, IEnumerable<MailAccount> mailboxes)
        {
            var itemView = new ItemView(int.MaxValue) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.HasAttachments) };

            Func<MailAccount, ExchangeService> ewsProvider = account => ImpersonateQueries(service, account.PrimarySmtpAddress);
            var serviceConfiguratorFor = ewsProvider.Memoize();

            var allFileAttachments = mailboxes
                .Select(acc => new { Mailbox = acc, Service = serviceConfiguratorFor(acc) })
                .SelectMany(x => GetAllCalendars(x.Service),
                    (x, Calendar) => new { x.Mailbox, x.Service, Calendar })
                .SelectMany(x => GetAppointmentsHavingAttachments(x.Calendar, itemView, x.Service),
                    (x, app) => new { Appointment = app, x.Mailbox, x.Calendar })
                .SelectMany(x => x.Appointment.Attachments.OfType<FileAttachment>(),
                    (x, att) => new AttachmentWithContext {
                        Mailbox = x.Mailbox,
                        Calendar = x.Calendar,
                        Appointment = x.Appointment,
                        Attachment = att,
                    })
                ;

            return allFileAttachments;
        }

        static IEnumerable<EWSAppointment> GetAppointmentsHavingAttachments(CalendarFolder calendar, ItemView itemView, ExchangeService service)
        {
            var events = calendar.FindItems(itemView).ToList();
            int nbAppointmentsWithAttachments = events.Count(e => e.HasAttachments);
            Console.WriteLine("found {0} appointments having attached files !", nbAppointmentsWithAttachments);
            var appointmentIdsHavingAttachments = events.Where(e => e.HasAttachments).Select(e => e.Id).ToList();
            if (appointmentIdsHavingAttachments.Count <= 0)
                return Enumerable.Empty<EWSAppointment>();
            var appointmentsWithAttachments = service.BindToItems(
                appointmentIdsHavingAttachments,
                new PropertySet(BasePropertySet.IdOnly,
                    ItemSchema.Attachments,
                    ItemSchema.HasAttachments,
                    ItemSchema.LastModifiedTime)
            );
            return appointmentsWithAttachments
                .Where(res => res.Result == ServiceResult.Success)
                .Select(res => res.Item)
                .Cast<EWSAppointment>();
        }

        private static void ExportAndPublishAddressBooks(MessageQueue queueConf, ExchangeService service, IEnumerable<MailAccount> mailboxes)
        {
            var folderView = new FolderView(100) {
                PropertySet = new PropertySet(
                    BasePropertySet.IdOnly,
                    FolderSchema.DisplayName,
                    FolderSchema.FolderClass),
                Traversal = FolderTraversal.Deep
            };
            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.FolderClass, "IPF.Contact");

            foreach(var box in mailboxes)
            {
                Console.WriteLine("Dumping CONTACTs for account: {0} ...", box.PrimarySmtpAddress);
                ImpersonateQueries(service, box.PrimarySmtpAddress);

                var rootFolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);

                var addressBooks = rootFolder.FindFolders(searchFilter, folderView);
                var addressBookMessages = addressBooks
                    .Select(book => new NewAddressBook
                    {
                        Id = Guid.NewGuid(),
                        CreationDate = DateTime.UtcNow,
                        PrimaryEmailAddress = box.PrimarySmtpAddress,
                        AddressBookId = book.Id.UniqueId,
                        DisplayName = book.DisplayName,
                    })
                    .ToList();
                addressBookMessages.ForEach(book => Console.WriteLine("Mailbox: {2}, Book #{0} , DisplayName: {1}", book.Id, book.DisplayName, book.PrimaryEmailAddress));

                PublishToBus(addressBookMessages, queueConf);
            }

        }

        private static void ExportAndPublishAppointments(MessageQueue queueConf, ExchangeService service, IEnumerable<MailAccount> mailboxes)
        {
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
        }

        private static ExchangeService ImpersonateQueries(ExchangeService service, string primaryAddress)
        {
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, primaryAddress);
            return service;
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

            var findAllAppointments = new Func<ExchangeService, FolderId,IEnumerable<EWSAppointment>>(FindAllAppointments).Partial(service);

            IQueryable<EWSAppointment> mailboxAppointments = GetAllCalendars(service)
                // .Where(cal => cal.DisplayName == "SubCalendar1" || cal.DisplayName == "SecondRootCalendar")
                .Select(calendar => calendar.Id)
                .SelectMany(findAllAppointments)
                .Cast<EWSAppointment>().AsQueryable();

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

        private static IEnumerable<EWSAppointment> FindAllAppointments(ExchangeService service, FolderId calendarId)
        {
            var appIdsView = new ItemView(int.MaxValue) {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType)
            };

            var result = PagedItemsSearch.PageSearchItems<EWSAppointment>(service, calendarId, 500, appIdsView.PropertySet, AppointmentSchema.DateTimeCreated);

            return result;
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
                            ItemSchema.Body,
                            AppointmentSchema.RequiredAttendees, 
                            AppointmentSchema.OptionalAttendees,
                            ItemSchema.Categories,
                            ItemSchema.Culture,
                            ItemSchema.DateTimeCreated,
                            ItemSchema.DateTimeReceived,
                            ItemSchema.DateTimeSent,
                            ItemSchema.DisplayTo,
                            ItemSchema.DisplayCc,
                            AppointmentSchema.Duration,
                            AppointmentSchema.End,
                            AppointmentSchema.Start,
                            AppointmentSchema.StartTimeZone,
                            ItemSchema.Subject,
                // AppointmentSchema.TextBody, // EWS 2013 only
                            AppointmentSchema.TimeZone,
                            ItemSchema.MimeContent,
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
            if (appointment.ModifiedOccurrences == null)
                return Convert(appointment);

            var reccurenceExceptionIds = appointment.ModifiedOccurrences.Select(occ => occ.ItemId);
            var reccurenceExceptionsAttendees = reccurenceExceptionIds
                .Select(id => EWSAppointment.Bind(service, id, new PropertySet( 
                    // TODO: Include Resources collection property
                    BasePropertySet.IdOnly, AppointmentSchema.RequiredAttendees, AppointmentSchema.OptionalAttendees)))
                .ToDictionary( k => ConvertIdFrom(k.Id), v => v );

            var dtoAppointment = Convert(appointment);

            // TODO: Consider adding OptionalAttendees and Resources as well in the dumped appointment DTO :)
            // TODO: include per attendee's iCal RSVP, found globally in appointment.IsResponseRequested property

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
