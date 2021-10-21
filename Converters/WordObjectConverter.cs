using System;
using JsonToWord.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonToWord.Converters
{
    public class WordObjectConverter : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            serializer.Serialize(writer, value, typeof(IWordObject));
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            var property = "type";
            var jo = JObject.Load(reader);

            if (string.Equals(jo[property].Value<string>(), WordObjectType.File.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<WordAttachment>(serializer);

            if (string.Equals(jo[property].Value<string>(), WordObjectType.Html.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<WordHtml>(serializer);

            if (string.Equals(jo[property].Value<string>(), WordObjectType.Paragraph.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<WordParagraph>(serializer);

            if (string.Equals(jo[property].Value<string>(), WordObjectType.Picture.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<WordAttachment>(serializer);

            if (string.Equals(jo[property].Value<string>(), WordObjectType.Table.ToString(), StringComparison.CurrentCultureIgnoreCase))
                return jo.ToObject<WordTable>(serializer);
            
            return serializer.Deserialize(reader, typeof(IWordObject));
        }

        public override bool CanConvert(Type objectType)
        {
            var result = (objectType == typeof(IWordObject));
            return result;
        }
    }
}