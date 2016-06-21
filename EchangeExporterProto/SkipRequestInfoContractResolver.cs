using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace EchangeExporterProto {

    public class SkipRequestInfoContractResolver : DefaultContractResolver {
        private readonly HashSet<string> excludedProperties;
        public SkipRequestInfoContractResolver(params string[] propertyNames) {
            excludedProperties = new HashSet<string>(propertyNames);
        }

        protected override JsonProperty CreateProperty(MemberInfo member, MemberSerialization memberSerialization) {
            if (excludedProperties.Any(s => s.Equals(member.Name, StringComparison.OrdinalIgnoreCase)))
                return default(JsonProperty);

            return base.CreateProperty(member, memberSerialization);
        }
    }
}
