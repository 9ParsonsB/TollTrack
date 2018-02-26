using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace TollTrack
{
    public partial class TrackingResult
    {
        [JsonProperty("key")]
        public string Key { get; set; }

        [JsonProperty("value")]
        public DateTime Value { get; set; }
    }

    public partial class TrackingResult
    {
    }
    public static class Converter
    {
        public static string ToJson<T>(this List<T> self) => JsonConvert.SerializeObject(self, Settings);
        public static List<T> FromJson<T>(this string json) => JsonConvert.DeserializeObject<List<T>>(json, Settings);

        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.DateTime,
            NullValueHandling = NullValueHandling.Ignore,
            Converters = { 
                new IsoDateTimeConverter
                {
                    DateTimeStyles = DateTimeStyles.AssumeUniversal,
                },
            },
        };
    }
}
