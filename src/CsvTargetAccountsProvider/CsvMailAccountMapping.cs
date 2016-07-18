using System;
using TinyCsvParser.Mapping;

namespace CsvTargetAccountsProvider
{
    public class CsvMailAccountMapping : CsvMapping<MailAccount>
    {
        public CsvMailAccountMapping()
            : base()
        {
            MapProperty(0, x => x.DistinguishedName);
            MapProperty(1, x => x.PrimarySmtpAddress);
            MapProperty(2, x => x.EmailAddresses);
        }
    }
}
