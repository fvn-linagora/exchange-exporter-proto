using System;
using System.Text;
using EasyNetQ;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace EchangeExporterProto {
    class NullHandingJsonSerializer : ISerializer
    {
        private readonly ITypeNameSerializer typeNameSerializer;

        private readonly JsonSerializerSettings serializerSettings = new JsonSerializerSettings
        {
            TypeNameHandling = TypeNameHandling.Auto,
            NullValueHandling = NullValueHandling.Ignore,
            Error = (serializer,err) => err.ErrorContext.Handled = true,
        };

        public NullHandingJsonSerializer(ITypeNameSerializer typeNameSerializer)
        {
            if (typeNameSerializer == null)
                throw new ArgumentNullException(nameof(typeNameSerializer));
            this.typeNameSerializer = typeNameSerializer;
            ConfigureEnumerationToBeSerializedAsString();
        }

        private void ConfigureEnumerationToBeSerializedAsString()
        {
            serializerSettings.Converters.Add(new StringEnumConverter { CamelCaseText = false });
        }

        public byte[] MessageToBytes<T>(T message) where T : class
        {
            if (message == null)
                throw new ArgumentNullException(nameof(message));
            return Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(message, serializerSettings));
        }

        public T BytesToMessage<T>(byte[] bytes)
        {
            if (bytes == null)
                throw new ArgumentNullException(nameof(bytes));
            return JsonConvert.DeserializeObject<T>(Encoding.UTF8.GetString(bytes), serializerSettings);
        }

        public object BytesToMessage(string typeName, byte[] bytes)
        {
            if (typeName == null)
                throw new ArgumentNullException(nameof(typeName));
            if (bytes == null)
                throw new ArgumentNullException(nameof(bytes));
            var type = typeNameSerializer.DeSerialize(typeName);
            return JsonConvert.DeserializeObject(Encoding.UTF8.GetString(bytes), type, serializerSettings);
        }
    }}
