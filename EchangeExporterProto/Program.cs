using System;
using System.Net;
using System.Linq;
using EasyNetQ;
using Messages;
using Microsoft.Exchange.WebServices.Data;

namespace EchangeExporterProto
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            service.Credentials = new WebCredentials("user1@MSLABLGS", "L1n4g0r4");
            service.Url = new Uri("https://172.16.24.101/EWS/Exchange.asmx");

            var fetchView = new ItemView(int.MaxValue);
            fetchView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties);

            //Newtonsoft.Json.JsonSerializer _jsonWriter = new Newtonsoft.Json.JsonSerializer {
            //    NullValueHandling = NullValueHandling.Ignore
            //};

            var queueConnectionString = "host=10.69.0.117";

            // using (var bus = RabbitHutch.CreateBus(queueConnectionString, RegisterSerializer(_jsonWriter)))
            using (var bus = RabbitHutch.CreateBus(queueConnectionString, serviceRegister => serviceRegister.Register<ISerializer>(
                    serviceProvider => new NullHandingJsonSerializer(new TypeNameSerializer()))))             
            {
                service.FindItems(WellKnownFolderName.Contacts, fetchView)
                    .Select(c => Contact.Bind(service, c.Id))
                    // .Cast<Contact>()
                    .ToList()
                    .ForEach(contact => bus.Publish(contact));
            }
            Console.ReadLine();
            // PublishMessage();
        }

        //private static Action<IServiceRegister> RegisterSerializer(Newtonsoft.Json.JsonSerializer _jsonWriter) {
        //    return serviceRegister => serviceRegister.Register<ISerializer>(serviceProvider => _jsonWriter);
        //}

        private static void PublishMessage() {
            using (var bus = RabbitHutch.CreateBus("host=10.69.0.117")) {
                var input = "";
                Console.WriteLine("Enter a message. 'Quit' to quit.");
                while ((input = Console.ReadLine()) != "Quit")
                    bus.Publish(new TextMessage {
                        Text = input
                    });
            }
        }
    }
}
