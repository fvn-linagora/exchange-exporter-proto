namespace EchangeDumpedMessagesListener
{
    using CommandLine;

    class ConfigurationOptions
    {
        [Option('c', "config",
           HelpText = "Configuration file path.")]
        public string ConfigPath { get; set; }
    }
}
