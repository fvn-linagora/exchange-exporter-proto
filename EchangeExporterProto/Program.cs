using System;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Exchange.WebServices.Data;
using EasyNetQ;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using Messages;
using SimpleConfig;

namespace EchangeExporterProto
{
    class Program
    {
        private static ExporterConfiguration config;

        private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service"),
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

            // Set default value when missing
            var queueConf = new MessageQueue {
                Host = config.MessageQueue.Host,
                Username = config.MessageQueue.Username ?? "guest",
                Password = config.MessageQueue.Password ?? "guest",
                VirtualHost = config.MessageQueue.VirtualHost ?? "/",
                Port = config.MessageQueue.Port != 0 ? config.MessageQueue.Port : 5672,
                ConnectionString = config.MessageQueue.ConnectionString
            };

            if (String.IsNullOrWhiteSpace(config.UserCredential.Domain)
                || String.IsNullOrWhiteSpace(config.UserCredential.Login)
                || String.IsNullOrWhiteSpace(config.UserCredential.Password))
            {
                Error("Provided credentials are incomplete!");
                return;
            }

            if (String.IsNullOrWhiteSpace(queueConf.ConnectionString))
                // config.MessageQueue.ConnectionString = String.Format("host={0}", config.MessageQueue.Host);
                queueConf.ConnectionString = String.Format("host={0};virtualHost={1};username={2};password={3}",
                    queueConf.Host, queueConf.VirtualHost, queueConf.Username, queueConf.Password);            

            ExchangeService service = ConnectToExchange(config.ExchangeServer, config.UserCredential);

            using (var bus = RabbitHutch.CreateBus(queueConf.ConnectionString , serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))
            {
                var foundEvents = FindAllMeetings(service);

                foreach (var ev in foundEvents)
                {
                    Console.WriteLine("Extracted with event #{0}. About to publish to {1}...", ev.Id, config.MessageQueue.Host);
                    bus.Publish(ev);
                }
            }
            Console.ReadLine();
        }

        private static string DumpAvailablePropsToJson(Appointment ev) {
            return JsonConvert.SerializeObject(ev, Formatting.Indented, serializerSettings);
        }

        private static IEnumerable<NewAppointmentDumped> FindAllMeetings(ExchangeService service)
        {
            PropertySet includeMostProps = BuildAppointmentPropertySet();
            var foundEvents = service.FindItems(WellKnownFolderName.Calendar, new ItemView(int.MaxValue));
            var foundEventsWithProps = foundEvents.Select(app => Appointment.Bind(service, app.Id, includeMostProps));

            var appointmentsInfo = foundEventsWithProps.Select(RemoveTypings);

            var appointmentAsMimeContent = foundEvents
                .Select(app => Appointment.Bind(service, app.Id, new PropertySet(AppointmentSchema.MimeContent)))
                .Select(app => app.MimeContent);

            foundEvents.Zip(appointmentsInfo, (app, dyn) => new NewAppointmentDumped {
                Id = app.Id.ToString(),
                Appointment = dyn,
            });

            return foundEvents
                .Zip(appointmentsInfo, (app, dyn) => new {
                    Id = app.Id.ToString(),
                    Appointment = dyn,
                })
                .Zip(appointmentAsMimeContent, (zip, mime) => new NewAppointmentDumped {
                    Id = zip.Id,
                    Appointment = zip.Appointment,
                    MimeContent = Encoding.GetEncoding(mime.CharacterSet).GetString(mime.Content),
                });
        }

        private static PropertySet BuildAppointmentPropertySet()
        {
            return new PropertySet(
                            BasePropertySet.FirstClassProperties,
                            AppointmentSchema.AppointmentType,
                            AppointmentSchema.Body,
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
                            AppointmentSchema.When
                        );
        }

        private static JObject RemoveTypings(Appointment app)
        {
            return JsonConvert.DeserializeObject<JObject>(
                JsonConvert.SerializeObject(app, Formatting.Indented, serializerSettings)
            );
        }

        private static ExchangeService ConnectToExchange(ExchangeServer exchange, UserCredential credential) {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            service.Credentials = new WebCredentials(String.Format("{0}@{1}", credential.Login, credential.Domain), credential.Password);
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
