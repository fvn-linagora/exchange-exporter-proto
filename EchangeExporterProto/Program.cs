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

namespace EchangeExporterProto
{
    class Program
    {
        private static readonly string EWSURL = "https://172.16.24.101/EWS/Exchange.asmx";
        private static readonly string EWSLOGIN = "user1@MSLABLGS";
        private static readonly string EWSPASS = "L1n4g0r4";
        private static readonly string MQCONNETIONSTRING = "host=10.69.0.117";

        private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service"),
            Error = (serializer, err) => err.ErrorContext.Handled = true,
        };

        static void Main(string[] args)
        {
            ExchangeService service = ConnectToExchange();

            using (var bus = RabbitHutch.CreateBus(MQCONNETIONSTRING, serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))
            {
                var foundEvents = FindAllMeetings(service);

                foreach (var ev in foundEvents)
                {
                    // Console.WriteLine(ev.Appointment.ToString());
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

        private static ExchangeService ConnectToExchange() {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            service.Credentials = new WebCredentials(EWSLOGIN, EWSPASS);
            service.Url = new Uri(EWSURL);
            return service;
        }
    }
}
