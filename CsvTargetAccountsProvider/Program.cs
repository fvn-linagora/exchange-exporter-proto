using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using TinyCsvParser;
using TinyCsvParser.Tokenizer.RegularExpressions;

namespace CsvTargetAccountsProvider
{
    class Program
    {
        private static readonly string ACCOUNTSFILE = "targets.csv";

        static void Main(string[] args)
        {
            var opt = new CsvParserOptions(true, new QuotedStringTokenizer(','));
            var mapper = new CsvMailAccountMapping();
            var parser = new CsvParser<MailAccount>(opt, mapper);

            var accountsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), ACCOUNTSFILE);

            var result = parser.ReadFromFile(accountsFilePath, Encoding.ASCII);

            result.Where(r => r.IsValid)
                .Select(acc => acc.Result.PrimarySmtpAddress).ToList()
                .ForEach(email => Console.WriteLine("dumping account: {0}", email));

            Console.ReadLine();
        }
    }
}
