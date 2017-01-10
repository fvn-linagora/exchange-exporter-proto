using System;
using System.Net;

namespace EchangeExporterProto
{
    public class ExporterConfiguration
    {
        public ExchangeServer ExchangeServer { get; set; }
        public MessageQueue MessageQueue { get; set; }
        public Credentials Credentials { get; set; }
    }

    public enum ExchangeServerVersion
    {
        Exchange2007_SP1,
        Exchange2010,
        Exchange2010_SP1,
        Exchange2010_SP2,
        Exchange2013,
        Exchange2013_SP1
    }

    public class ExchangeServer
    {
        public string EndpointTemplate { get; set; }
        public string Host { get; set; }
        public ExchangeServerVersion ExchangeVersion { get; set; } = ExchangeServerVersion.Exchange2010_SP2;
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

    public class Credentials
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
    }
}
