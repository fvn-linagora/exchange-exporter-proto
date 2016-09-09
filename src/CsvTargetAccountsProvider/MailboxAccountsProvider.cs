namespace CsvTargetAccountsProvider
{
    using System.Linq;
    using System.Collections.Generic;
    using System.Text;
    using TinyCsvParser;
    using TinyCsvParser.Tokenizer.RegularExpressions;

    public class MailboxAccountsProvider
    {
        private readonly CsvParserOptions opt;
        private readonly CsvMailAccountMapping mapper = new CsvMailAccountMapping();
        private readonly CsvParser<MailAccount> parser;

        public MailboxAccountsProvider(char delimiter)
        {
            opt = new CsvParserOptions(true, new QuotedStringTokenizer(delimiter));
            parser = new CsvParser<MailAccount>(opt, mapper);
        }

        public IEnumerable<MailAccount> GetFromCsvFile(string inputCSV)
        {
            var result = parser.ReadFromFile(inputCSV, Encoding.ASCII);
            return result.Where(r => r.IsValid)
                .Select(acc => acc.Result);
        }
    }
}
