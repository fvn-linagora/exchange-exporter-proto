using System;
using System.Net;

namespace EchangeExporterProto
{
    public class ExporterConfiguration
    {
        public ExchangeServer ExchangeServer { get; set; }
        public MessageQueue MessageQueue { get; set; }
        public UserCredential UserCredential { get; set; }
    }

    public class ExchangeServer
    {
        public string EndpointTemplate { get; set; }
        public string Host { get; set; }
    }

    public class MessageQueue
    {
        public string ConnectionString { get; set; }
        public string Host { get; set; }
        public string VirtualHost { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    public class UserCredential
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
    }
}
