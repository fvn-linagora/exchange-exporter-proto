using System;
using System.Text;
using EasyNetQ;
using Newtonsoft.Json;

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
                throw new ArgumentNullException("typeNameSerializer");
            this.typeNameSerializer = typeNameSerializer;
        }

        public byte[] MessageToBytes<T>(T message) where T : class
        {
            if (message == null)
                throw new ArgumentNullException("message");
            return Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(message, serializerSettings));
        }

        public T BytesToMessage<T>(byte[] bytes)
        {
            if (bytes == null)
                throw new ArgumentNullException("bytes");
            return JsonConvert.DeserializeObject<T>(Encoding.UTF8.GetString(bytes), serializerSettings);
        }

        public object BytesToMessage(string typeName, byte[] bytes)
        {
            if (typeName == null)
                throw new ArgumentNullException("typeName");
            if (bytes == null)
                throw new ArgumentNullException("bytes");
            var type = typeNameSerializer.DeSerialize(typeName);
            return JsonConvert.DeserializeObject(Encoding.UTF8.GetString(bytes), type, serializerSettings);
        }
    }}
