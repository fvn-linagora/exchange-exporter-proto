using System;
using System.IO;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Exchange.WebServices.Data;
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

        private static string DumpAvailablePropsToJson(Appointment ev) {
            return JsonConvert.SerializeObject(ev, Formatting.Indented, serializerSettings);
        }

        private static IEnumerable<NewAppointmentDumped> FindAllMeetings(ExchangeService service, String primaryEmailAddress)
        {
            PropertySet includeMostProps = BuildAppointmentPropertySet();
            var appIdsView = new ItemView(int.MaxValue) { PropertySet = BasePropertySet.IdOnly};

            var foundEventsWithProps = GetAllCalendars(service)
                // .Where(cal => cal.DisplayName == "SubCalendar1" || cal.DisplayName == "SecondRootCalendar")
                .SelectMany(calendar => service.FindItems(calendar.Id, appIdsView))
                .Select(app => Appointment.Bind(service, app.Id, includeMostProps))
                .Select(app => new {
                    Mailbox = primaryEmailAddress,
                    Folder = app.ParentFolderId,
                    Appointment = app
                });

            var foundEventsAsMime = foundEventsWithProps.Select(appCtx => appCtx.Appointment.MimeContent);

            return foundEventsWithProps
                .Select(appCtx => new NewAppointmentDumped {
                    Mailbox = appCtx.Mailbox,
                    Id = appCtx.Appointment.Id.ToString(),
                    Appointment = RemoveTypings(appCtx.Appointment),
                    MimeContent = Encoding.GetEncoding(appCtx.Appointment.MimeContent.CharacterSet).GetString(appCtx.Appointment.MimeContent.Content)
                });
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
                            AppointmentSchema.When
                        );
        }

        private static JObject RemoveTypings(Appointment app)
        {
            return JsonConvert.DeserializeObject<JObject>(
                JsonConvert.SerializeObject(app, Formatting.Indented, serializerSettings)
            );
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
