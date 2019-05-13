using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public partial class SmartsheetShare
    {
        [JsonProperty("pageNumber", NullValueHandling = NullValueHandling.Ignore)]
        public long? PageNumber { get; set; }

        [JsonProperty("pageSize", NullValueHandling = NullValueHandling.Ignore)]
        public long? PageSize { get; set; }

        [JsonProperty("totalPages", NullValueHandling = NullValueHandling.Ignore)]
        public long? TotalPages { get; set; }

        [JsonProperty("totalCount", NullValueHandling = NullValueHandling.Ignore)]
        public long? TotalCount { get; set; }

        [JsonProperty("data", NullValueHandling = NullValueHandling.Ignore)]
        public List<Datum> Data { get; set; }
    }

    public partial class Datum
    {
        [JsonProperty("id", NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty("type", NullValueHandling = NullValueHandling.Ignore)]
        public string Type { get; set; }

        [JsonProperty("userId", NullValueHandling = NullValueHandling.Ignore)]
        public long? UserId { get; set; }

        [JsonProperty("email", NullValueHandling = NullValueHandling.Ignore)]
        public string Email { get; set; }

        [JsonProperty("name", NullValueHandling = NullValueHandling.Ignore)]
        public string Name { get; set; }

        [JsonProperty("accessLevel", NullValueHandling = NullValueHandling.Ignore)]
        public string AccessLevel { get; set; }

        [JsonProperty("scope", NullValueHandling = NullValueHandling.Ignore)]
        public string Scope { get; set; }

        [JsonProperty("createdAt", NullValueHandling = NullValueHandling.Ignore)]
        public DateTimeOffset? CreatedAt { get; set; }

        [JsonProperty("modifiedAt", NullValueHandling = NullValueHandling.Ignore)]
        public DateTimeOffset? ModifiedAt { get; set; }

        [JsonProperty("groupId", NullValueHandling = NullValueHandling.Ignore)]
        public long? GroupId { get; set; }
    }

    public partial class SmartsheetShare
    {
        public static SmartsheetShare FromJson(string json) => JsonConvert.DeserializeObject<SmartsheetShare>(json, MicrosoftGraphAspNetCoreConnectSample.Helpers.Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this SmartsheetShare self) => JsonConvert.SerializeObject(self, MicrosoftGraphAspNetCoreConnectSample.Helpers.Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }
}
