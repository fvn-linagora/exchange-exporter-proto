using System.Collections.Generic;
using CommandLine;

namespace EchangeExporterProto
{
    class ExporterOptions
    {
        [Option('t', "targets", Required = true,
             HelpText = "Input mailboxes to be processed.")]
        public string TargetsListFile { get; set; }

        [Option('c', "config", Required = true,
             HelpText = "Configuration file path.")]
        public string ConfigPath { get; set; }

        [Option('s', "skip-steps", HelpText = "Steps to be skipped (events, addressbooks, attachments, contacts).")]
        public IEnumerable<Features> SkippedSteps { get; set; }

        [Option(
             HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }
    }
}