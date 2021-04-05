using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Aspose.Slides.Web.API.Json
{
	public class BooleanNullHandlingConverter : JsonConverter<bool>
	{
		public override bool Read(
			ref Utf8JsonReader reader,
			Type typeToConvert,
			JsonSerializerOptions options) =>
			reader.TokenType == JsonTokenType.Null
				? default
				: reader.GetBoolean();

		public override void Write(
			Utf8JsonWriter writer,
			bool boolValue,
			JsonSerializerOptions options) =>
			writer.WriteBooleanValue(boolValue);
	}
}
