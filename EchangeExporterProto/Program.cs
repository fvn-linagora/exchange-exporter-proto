using System;
using System.Net;
using System.Linq;
using EasyNetQ;
using Microsoft.Exchange.WebServices.Data;
using System.Text;
using Messages;
using System.Collections.Generic;
using Newtonsoft.Json;

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

            var fetchView = new ItemView(int.MaxValue);
            fetchView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);

            using (var bus = RabbitHutch.CreateBus(MQCONNETIONSTRING, serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))             
            {
                var foundEvents = FindAllMeetings(service);

                foreach (var ev in foundEvents)
                {
                    // Console.WriteLine(ev);
                    Console.WriteLine(DumpAvailablePropsToJson(ev));
                    bus.Publish(ev);
                }
            }
            Console.ReadLine();
        }

        private static string DumpAvailablePropsToJson(Appointment ev) {
            return JsonConvert.SerializeObject(ev, Formatting.Indented, serializerSettings);
        }

        private static IEnumerable<Appointment> FindAllMeetings(ExchangeService service)
        {
            //service.FindItems(WellKnownFolderName.Contacts, fetchView)
            //    .Select(c => Contact.Bind(service, c.Id))
            //    .ToList()
            //    .ForEach(bus.Publish);

            // PropertySet includeMimeICS = new PropertySet(BasePropertySet.FirstClassProperties, AppointmentSchema.MimeContent);
            PropertySet includeMostProps = new PropertySet(
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
            var foundEvents = service.FindItems(WellKnownFolderName.Calendar, new ItemView(int.MaxValue))
                .Select(app => Appointment.Bind(service, app.Id, includeMostProps))
                // .Select(app => Appointment.Bind(service, app.Id))
                // .Select(app => Encoding.GetEncoding(app.MimeContent.CharacterSet).GetString(app.MimeContent.Content))
                // .Select(DumpAsEvent)
                // .Select(DumpRequestMetadata)
                ;

            return foundEvents;
        }

        private static IEvent DumpAsEvent(Appointment app)
        {
            MimeContent appointmentMimeContent = new MimeContent(Encoding.UTF8.WebName, new byte[] {});
            try {
                appointmentMimeContent = app.MimeContent;
            }
            catch {}

            return new NewAppointmentDumped {
                Id = app.Id.ToString(),
                // Appointment = app,
                MimeContent = Encoding.GetEncoding(appointmentMimeContent.CharacterSet).GetString(appointmentMimeContent.Content),
            };
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
