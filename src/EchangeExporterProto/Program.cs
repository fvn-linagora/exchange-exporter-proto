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
using Appointment = Messages.Appointment;

namespace EchangeExporterProto
{
    enum Features
    {
        Event,
        AddressBook,
        Attachment,
        Contact
    }

    class Options
    {
        // [Option('t', "targets", Required = true,
        [Option('t', "targets",
          HelpText = "Input mailboxes to be processed.")]
        public string TargetsListFile { get; set; }

        [Option('c', "config",
           HelpText = "Configuration file path.")]
        public string ConfigPath { get; set; }

        [Option('s', "skip-steps", HelpText = "Steps to be skipped (events, addressbooks, attachments, contacts).")]
        public IEnumerable<Features> SkippedSteps { get; set; }

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
        private static ISet<Features> skippedSteps = new HashSet<Features>();

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
            var mailboxes = GetTargetAccounts(arguments)
                .Where(box => !string.IsNullOrWhiteSpace(box.PrimarySmtpAddress))
                .ToList();
            skippedSteps = new HashSet<Features>(arguments.SkippedSteps);

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

            if (!skippedSteps.Contains(Features.Event))
                ExportAndPublishAppointments(queueConf, service, mailboxes);
            if (!(skippedSteps.Contains(Features.AddressBook) && skippedSteps.Contains(Features.Contact)))
                ExportAndPublishAddressBooks(queueConf, service, mailboxes);

            if (!skippedSteps.Contains(Features.Attachment))
            {
                var attachedMessages = ExportAppointmentsAttachedFiles(service, mailboxes).Select(MapToAttachmentMessage);
                PublishToBus(attachedMessages, queueConf);
            }

            Console.WriteLine("DONE: Exporter has completed exchange mailboxes data dump.");
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

            foreach (var box in mailboxes)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(box.PrimarySmtpAddress)) continue;

                    Console.WriteLine("Dumping Addressbooks for account: {0} ...", box.PrimarySmtpAddress);
                    ImpersonateQueries(service, box.PrimarySmtpAddress);

                    var rootFolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);

                    var addressBooks = rootFolder.FindFolders(searchFilter, folderView).Cast<ContactsFolder>().ToList();

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

                    if (!skippedSteps.Contains(Features.Contact))
                    {
                        IEnumerable<NewMimeContactExported> contactMessages = DumpAddressBookContacts(service, box.PrimarySmtpAddress, addressBooks.Cast<ContactsFolder>());
                        PublishToBus(contactMessages, queueConf);
                    }
                }
                catch (Exception ex)
                {
                    log.Error($"An error occured for mailbox {box.PrimarySmtpAddress}, message: {ex.Message}, stack: {ex.StackTrace}");
                }
            }
        }

        private static IEnumerable<NewMimeContactExported> DumpAddressBookContacts(ExchangeService service, String primaryAddress, IEnumerable<ContactsFolder> addressBooks)
        {
            var includingOnlyIdAndName = new ItemView(int.MaxValue) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ContactSchema.DisplayName) };
            var includingMimeAndLastUpdated = new PropertySet(BasePropertySet.IdOnly, ItemSchema.MimeContent, ContactSchema.DisplayName, ItemSchema.LastModifiedTime);
            var allBookIdsWithContactIds = addressBooks
                .Select(book => new {
                    BookId = book.Id.UniqueId,
                    ContactIds = book.FindItems(includingOnlyIdAndName).Select(x => x.Id),
                })
                .Where(x => x.ContactIds.Any());

            var allContactsInfo = allBookIdsWithContactIds
                .SelectMany(x => service.BindToItems(x.ContactIds, includingMimeAndLastUpdated)
                    .Select(resp => new { x.BookId, Response = resp }))
                .Where(x => AddressBookItemResponseIsSingleContact(x.Response)) // skip Distribution Lists
                .Select(x => new ContactContext {
                    PrimaryAddress = primaryAddress,
                    AddressBookId = x.BookId,
                    Contact = x.Response.Item as Contact
                });

            Func<Contact, DateTime> getLastModifiedUtc = contact => TimeZoneInfo.ConvertTimeToUtc(contact.LastModifiedTime, service.TimeZone);
            Func<MimeContent, string> mimeToString = mime => Encoding.GetEncoding(mime.CharacterSet).GetString(mime.Content);
            Func<ContactContext, NewMimeContactExported> createExportedContactMessage = contactContext => CreateExportedContactMessage(getLastModifiedUtc, mimeToString, contactContext);

            return allContactsInfo
                .Select(createExportedContactMessage);
        }

        private static bool AddressBookItemResponseIsSingleContact(GetItemResponse contactResponse)
        {
            return contactResponse.Result == ServiceResult.Success
                    && contactResponse.Item is Contact
                    && (contactResponse.Item as Contact) != null;
        }

        class ContactContext
        {
            public string PrimaryAddress { get; set; }
            public string AddressBookId { get; set; }
            public Contact Contact { get; set; }
        }

        private static NewMimeContactExported CreateExportedContactMessage(Func<Contact, DateTime> getLastModifiedUtc, Func<MimeContent, string> mimeToString, ContactContext ctx)
        {
            if (mimeToString == null)
                throw new ArgumentNullException(nameof(mimeToString));
            if (getLastModifiedUtc == null)
                throw new ArgumentNullException(nameof(getLastModifiedUtc));
            if (ctx == null)
                throw new ArgumentNullException(nameof(ctx));
            if (ctx.Contact == null)
                throw new ArgumentNullException(nameof(ctx.Contact));
            if (string.IsNullOrWhiteSpace(ctx.PrimaryAddress))
                throw new ArgumentNullException(nameof(ctx.PrimaryAddress));
            if (ctx.AddressBookId == null)
                throw new ArgumentNullException(nameof(ctx.AddressBookId));
            if (ctx.Contact.MimeContent == null)
                throw new ArgumentNullException(nameof(ctx.Contact.MimeContent));
            if (ctx.Contact.Id == null || ctx.Contact.Id.UniqueId == null || ctx.Contact.Id.ChangeKey == null)
                throw new ArgumentNullException(nameof(ctx.Contact.Id));

            log.Info($"exporting contact '{ctx.Contact.DisplayName ?? string.Empty}' from mailbox: '{ctx.PrimaryAddress}'\n");

            return new NewMimeContactExported
            {
                Id = Guid.NewGuid(),
                CreationDate = DateTimeOffset.UtcNow,
                OriginalChangeKey = ctx.Contact.Id.ChangeKey,
                LastModified = getLastModifiedUtc(ctx.Contact),
                OriginalContactId = ctx.Contact.Id.UniqueId,
                AddressBookId = ctx.AddressBookId,
                PrimaryAddress = ctx.PrimaryAddress,
                MimeContent = mimeToString(ctx.Contact.MimeContent),
            };
        }

        private static void ExportAndPublishAppointments(MessageQueue queueConf, ExchangeService service, IEnumerable<MailAccount> mailboxes)
        {
            using (var bus = RabbitHutch.CreateBus(queueConf.ConnectionString , serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))
            {
                foreach (var mailbox in mailboxes)
                {
                    if (string.IsNullOrWhiteSpace(mailbox.PrimarySmtpAddress)) continue;
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

        class TractableJsonSerializer
        {
            private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings
            {
                TypeNameHandling = TypeNameHandling.Auto,
                NullValueHandling = NullValueHandling.Ignore,
                ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service", "MimeContent"),
                Error = (serializer, err) => err.ErrorContext.Handled = true,
            };

            public string ToJson(object value)
            {
                return JsonConvert.SerializeObject(value, Formatting.Indented, serializerSettings);
            }
        }

        class ContextualizedAppointment
        {
            public string PrimarySmtpAddress { get; set; }

            public Messages.ItemId Folder { get; set; }

            public Messages.Appointment Appointment { get; set; }
        }

        interface IAppointmentsProvider
        {
            IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress);
        }

        class ExchangeAppointmentsProvider : IAppointmentsProvider
        {
            private readonly ExchangeService service;
            private readonly TractableJsonSerializer serializer;

            public IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress)
            {

            }

            private IEnumerable<Appointment> FindAllMeetings(string primaryEmailAddress)
            {
                PropertySet includeMostProps = BuildAppointmentPropertySet();

                var findAllAppointments = new Func<ExchangeService, FolderId, IEnumerable<EWSAppointment>>(FindAllAppointments)
                    .Partial(service);

                var mailboxAppointments = GetAllCalendars(service)
                    .Select(calendar => calendar.Id)
                    .SelectMany(findAllAppointments);

                var singleAndReccurringMasterAppointments = mailboxAppointments
                    .Where(app => singleAndRecurringMasterAppointmentTypes.Contains(app.AppointmentType));

                var singleAndReccurringMasterAppointmentsWithContext = singleAndReccurringMasterAppointments
                    .Select(app => EWSAppointment.Bind(service, app.Id, includeMostProps))
                    .Select(appointment => AddMissingAttendeesInfo(service, appointment));
//                    .Select(app => new ContextualizedAppointment {
//                        PrimarySmtpAddress = primaryEmailAddress,
//                        Folder = ConvertIdFrom(app.ParentFolderId),
//                        Appointment = Convert(app)
//                    });

                return singleAndReccurringMasterAppointmentsWithContext;

//                var messagesForExportingSingleAndReccurenceAppointments = singleAndReccurringMasterAppointmentsWithContext
//                    .Select(appCtx => new NewAppointmentDumped
//                    {
//                        Mailbox = appCtx.Mailbox,
//                        FolderId = appCtx.Folder.UniqueId,
//                        Id = appCtx.Appointment.Id.ToString(),
//                        Appointment = AddMissingAttendeesInfo(service, appCtx.Appointment),
//                        SourceAsJson = serializer.ToJson(appCtx.Appointment),
//                        MimeContent = Encoding.GetEncoding(appCtx.Appointment.MimeContent.CharacterSet)
//                            .GetString(appCtx.Appointment.MimeContent.Content)
//                    });
//
//                return messagesForExportingSingleAndReccurenceAppointments;
            }

            private static Messages.Appointment Convert(EWSAppointment app)
            {
                // TODO: there should be a genuine mapping eventually here, as models may diverge !
                var serializedPayload = JsonConvert.SerializeObject(app, Formatting.Indented, serializerSettings);
                return JsonConvert.DeserializeObject<Messages.Appointment>(serializedPayload);
            }

            private static RequiredAttendee MapToRequiredAttendees(Microsoft.Exchange.WebServices.Data.Attendee att)
            {
                return new RequiredAttendee
                {
                    Address = att.Address,
                    Name = att.Name,
                    MailboxType = (Messages.MailboxType) (int?) att.MailboxType,
                    ResponseType = (Messages.MeetingResponseType) (int?) att.ResponseType,
                    RoutingType = att.RoutingType,
                };
            }

            private static OptionalAttendee MapToOptionalAttendees(Microsoft.Exchange.WebServices.Data.Attendee att)
            {
                return new OptionalAttendee
                {
                    Address = att.Address,
                    Name = att.Name,
                    MailboxType = (Messages.MailboxType) (int?) att.MailboxType,
                    ResponseType = (Messages.MeetingResponseType) (int?) att.ResponseType,
                    RoutingType = att.RoutingType,
                };
            }

            private static Resource MapToResources(Microsoft.Exchange.WebServices.Data.Attendee att)
            {
                return new Resource
                {
                    Address = att.Address,
                    Name = att.Name,
                    MailboxType = (Messages.MailboxType) (int?) att.MailboxType,
                    ResponseType = (Messages.MeetingResponseType) (int?) att.ResponseType,
                    RoutingType = att.RoutingType,
                };
            }

            private static Messages.ItemId ConvertIdFrom(ServiceId id)
            {
                return new Messages.ItemId
                {
                    UniqueId = id.UniqueId,
                    ChangeKey = id.ChangeKey
                };
            }

            private static IEnumerable<EWSAppointment> FindAllAppointments(ExchangeService service, FolderId calendarId)
            {
                var appIdsView = new ItemView(int.MaxValue)
                {
                    PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType)
                };

                var result = PagedItemsSearch.PageSearchItems<EWSAppointment>(service, calendarId, 500,
                    appIdsView.PropertySet, AppointmentSchema.DateTimeCreated);

                return result;
            }

            private static IEnumerable<CalendarFolder> GetAllCalendars(ExchangeService service)
            {
                int folderViewSize = int.MaxValue; // Kids, dont do this at home !
                var view = new FolderView(folderViewSize)
                {
                    PropertySet =
                        new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.FolderClass),
                    Traversal = FolderTraversal.Deep
                };
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
                    AppointmentSchema.Resources,
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
                    AppointmentSchema.TimeZone,
                    ItemSchema.MimeContent,
                    AppointmentSchema.ModifiedOccurrences,
                    AppointmentSchema.DeletedOccurrences,
                    AppointmentSchema.AppointmentType,
                    AppointmentSchema.IsResponseRequested,
                    AppointmentSchema.When
                );
            }

        }

        private static ExchangeService ConnectToExchange(ExchangeServer exchange, Credentials credentials) {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            // Ignoring invalid exchange server provided certificate, on purpose, Yay !
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;

            service.Credentials = new NetworkCredential(credentials.Login, credentials.Password, credentials.Domain);
            string exchangeEndpoint = string.Format(exchange.EndpointTemplate, exchange.Host);
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
