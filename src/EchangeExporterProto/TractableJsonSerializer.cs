using Newtonsoft.Json;

namespace EchangeExporterProto
{
    public class TractableJsonSerializer
    {
        private static readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings
        {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new SkipRequestInfoContractResolver("Schema", "Service", "MimeContent"),
            Error = (serializer, err) => err.ErrorContext.Handled = true,
        };

        public string ToJson(object value)
        {
            return JsonConvert.SerializeObject(value, Formatting.Indented, serializerSettings);
        }
    }
}