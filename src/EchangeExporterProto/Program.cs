using System;
using System.IO;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Exchange.WebServices.Data;
using EWSAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
using EasyNetQ;
using SimpleConfig;
using CommandLine;

using Messages;
using CsvTargetAccountsProvider;

namespace EchangeExporterProto
{
    enum Features
    {
        Event,
        AddressBook,
        Attachment,
        Contact
    }

    class Program
    {
        private const string EXPORTER_CONFIG_SECTION = "exporterConfiguration";
        private const string DEFAULT_EXPORTER_CONFIG = "ExchangeExporter.config";
        private const string ENV_EXPORTER_CONFIG = "EXPORTER_CONFIG";
        private const string ACCOUNTSFILE = "targets.csv";
        private static ExporterConfiguration config;
        private static readonly MailboxAccountsProvider accountsProvider = new MailboxAccountsProvider(',');
        private static ISet<Features> skippedSteps = new HashSet<Features>();

        private static readonly ILog log = new ConsoleLogger();
        private static readonly TractableJsonSerializer serializer = new TractableJsonSerializer();
        private static IAppointmentsProvider appointmentsProvider;
        private static readonly ISet<string> suggestedContactsFolderName = new SortedSet<string>(new string[] { @"Suggested Contacts", @"Contacts suggérés" });

        static void Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<ExporterOptions>(args);
            var arguments = result.MapResult( options => options, ArgumentErrorHandler);

            config = new Configuration(configPath: GetConfigPath(arguments)).LoadSection<ExporterConfiguration>(EXPORTER_CONFIG_SECTION);
            var mailboxes = GetTargetAccounts(arguments)
                .Where(box => !string.IsNullOrWhiteSpace(box.PrimarySmtpAddress))
                .ToList();
            skippedSteps = new HashSet<Features>(arguments.SkippedSteps);

            if (skippedSteps.Count == Enum.GetValues(typeof(Features)).Length)
                log.Warn($"Note that you chose to skip all steps ! Application won't do much at this point :) ");

            if (string.IsNullOrWhiteSpace(config.MessageQueue.ConnectionString) && string.IsNullOrWhiteSpace(config.MessageQueue.Host))
            {
                Error("Could not find either a connection string or an host for MQ!");
                return;
            }

            var queueConf = CompleteQueueConfigWithDefaults();

            if (string.IsNullOrWhiteSpace(config.Credentials.Domain)
                || string.IsNullOrWhiteSpace(config.Credentials.Login)
                || string.IsNullOrWhiteSpace(config.Credentials.Password))
            {
                Error("Provided credentials are incomplete!");
                return;
            }

            if (string.IsNullOrWhiteSpace(queueConf.ConnectionString))
                queueConf.ConnectionString =
                    $"host={queueConf.Host};virtualHost={queueConf.VirtualHost};username={queueConf.Username};password={queueConf.Password}";

            ExchangeService service = ConnectToExchange(config.ExchangeServer, config.Credentials);
            appointmentsProvider = new ExchangeAppointmentsProvider(serializer, pa => ImpersonateQueries(service, pa));

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

        private static ExporterOptions ArgumentErrorHandler(IEnumerable<Error> errors) {
            log.Error($"Found issues with '{string.Join("\n", errors)}'");
            Environment.Exit(1);
            return default(ExporterOptions);
        }

        private static string GetConfigPath(ExporterOptions arguments)
        {
            string configPath =
                // First check args
                (string.IsNullOrWhiteSpace(arguments.ConfigPath) ? null : arguments.ConfigPath)
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

        private static IEnumerable<MailAccount> GetTargetAccounts(ExporterOptions arguments)
        {
            var accountsFilePath = FindTargetAccountsFile(arguments) ??
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ACCOUNTSFILE);
            return accountsProvider.GetFromCsvFile(accountsFilePath);
        }

        private static string FindTargetAccountsFile(ExporterOptions arguments)
        {
            if (string.IsNullOrWhiteSpace(arguments.TargetsListFile))
                return null;
            return File.Exists(arguments.TargetsListFile) ?
                arguments.TargetsListFile : null;
        }

        private static IEnumerable<AttachmentWithContext> ExportAppointmentsAttachedFiles(ExchangeService service, IEnumerable<MailAccount> mailboxes)
        {
            var itemView = new ItemView(int.MaxValue) { PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.HasAttachments) };

            Func<MailAccount, ExchangeService> ewsProvider = account => ImpersonateQueries(service, account.PrimarySmtpAddress);
            var serviceConfiguratorFor = ewsProvider.Memoize();

            var allFileAttachments = mailboxes
                .Select(acc => new { Mailbox = acc, Service = serviceConfiguratorFor(acc) })
                .SelectMany(x => ExchangeAppointmentsProvider.GetAllCalendars(x.Service),
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

        private static IEnumerable<EWSAppointment> GetAppointmentsHavingAttachments(Folder calendar, ItemView itemView, ExchangeService service)
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

                    var defaultAddressBook = Folder.Bind(service, WellKnownFolderName.Contacts) as ContactsFolder;
                    var rootFolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot);
                    var addressBooks = rootFolder.FindFolders(searchFilter, folderView).Cast<ContactsFolder>().ToList();

                    var addressBookMessages = addressBooks
                        .Select(book => new NewAddressBook
                        {
                            Id = Guid.NewGuid(),
                            CreationDate = DateTime.UtcNow,
                            PrimaryEmailAddress = box.PrimarySmtpAddress,
                            AddressBookId = book.Id.UniqueId,
                            AddressBookType = GetAddressBookType(book, defaultAddressBook),
                            DisplayName = book.DisplayName,
                        })
                        .ToList();
                    addressBookMessages.ForEach(book => Console.WriteLine("Mailbox: {2}, Book #{0} , DisplayName: {1}", book.Id, book.DisplayName, book.PrimaryEmailAddress));
                    PublishToBus(addressBookMessages, queueConf);

                    if (skippedSteps.Contains(Features.Contact)) continue;

                    IEnumerable<NewMimeContactExported> contactMessages = DumpAddressBookContacts(service, box.PrimarySmtpAddress, addressBooks.Cast<ContactsFolder>());
                    PublishToBus(contactMessages, queueConf);
                }
                catch (Exception ex)
                {
                    log.Error($"An error occured for mailbox {box.PrimarySmtpAddress}, message: {ex.Message}, stack: {ex.StackTrace}");
                }
            }
        }

        private static AddressBookType GetAddressBookType(ContactsFolder book, ContactsFolder defaultAddressBook)
        {
            if (defaultAddressBook == null || string.IsNullOrWhiteSpace(defaultAddressBook.Id?.UniqueId))
                throw new ArgumentNullException(nameof(defaultAddressBook));
            if (book == null || string.IsNullOrWhiteSpace(book.Id?.UniqueId))
                throw new ArgumentNullException(nameof(book));

            if (book.Id.UniqueId == defaultAddressBook.Id.UniqueId)
                return AddressBookType.Primary;
            if (suggestedContactsFolderName.Contains(book?.DisplayName, StringComparer.InvariantCultureIgnoreCase))
                return AddressBookType.Collected;
            return AddressBookType.Custom;
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
            Func<Microsoft.Exchange.WebServices.Data.MimeContent, string> mimeToString = mime => Encoding.GetEncoding(mime.CharacterSet).GetString(mime.Content);
            Func<ContactContext, NewMimeContactExported> createExportedContactMessage = contactContext => CreateExportedContactMessage(getLastModifiedUtc, mimeToString, contactContext);

            return allContactsInfo
                .Select(createExportedContactMessage);
        }

        private static bool AddressBookItemResponseIsSingleContact(GetItemResponse contactResponse)
        {
            return contactResponse.Result == ServiceResult.Success && contactResponse.Item as Contact != null;
        }

        class ContactContext
        {
            public string PrimaryAddress { get; set; }
            public string AddressBookId { get; set; }
            public Contact Contact { get; set; }
        }

        private static NewMimeContactExported CreateExportedContactMessage(Func<Contact, DateTime> getLastModifiedUtc,
            Func<Microsoft.Exchange.WebServices.Data.MimeContent, string> mimeToString, ContactContext ctx)
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
            if (ctx.Contact.Id?.UniqueId == null || ctx.Contact.Id.ChangeKey == null)
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

                    // NOTE: as exchangeservice instance is shared & mutated, provider is inherently NOT thread-safe !
                    var appointments = appointmentsProvider.FindByMailbox(mailbox.PrimarySmtpAddress);

                    var foundEvents = appointments.Select(app => new NewAppointmentDumped
                    {
                        Mailbox = mailbox.PrimarySmtpAddress,
                        FolderId = app.ParentFolderId.UniqueId,
                        Id = app.Id.UniqueId,
                        Appointment = app,
                        SourceAsJson = serializer.ToJson(app),
                        MimeContent = app.AsICal
                    });

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
