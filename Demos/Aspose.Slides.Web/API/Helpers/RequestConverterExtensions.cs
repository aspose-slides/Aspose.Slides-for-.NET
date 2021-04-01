using System;
using Aspose.Slides.Web.API.Clients.DTO;
using Aspose.Slides.Web.API.Clients.DTO.Request;
using Aspose.Slides.Web.Interfaces.Models.Watermark;
using Aspose.Slides.Web.Interfaces.Models.Redaction;
using Aspose.Slides.Web.Interfaces.Models.Metadata;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace Aspose.Slides.Web.API.Helpers
{
	internal static class RequestConverterExtensions
	{
		/// <summary>
		/// Converts ImageWatermarkOptionsModel to ImageWatermarkOptions
		/// </summary>
		/// <returns></returns>
		public static ImageWatermarkOptions GetOptions(this ImageWatermarkOptionsRequest request)
		{
			var options = new ImageWatermarkOptions
			{
				idMain = request.idMain,
				MainFileNames = new List<string> ( request.MainFileNames.Select(s => s)),
				IsGrayScaled = request.IsGrayScaled,
				ZoomPercent = request.ZoomPercent,
				RotationAngleDegrees = request.RotationAngleDegrees
			};

			options.SetImageFile(request.ImageFile);

			return options;
		}

		/// <summary>
		/// Converts TextWatermarkOptionsModel to TextWatermarkOptions
		/// </summary>
		/// <returns></returns>
		public static TextWatermarkOptions GetOptions(this TextWatermarkOptionsRequest request)
		{
			return new TextWatermarkOptions
			{
				Text = request.Text,
				Color = request.Color,
				FontName = request.FontName,
				FontSize = request.FontSize,
				RotationAngleDegrees = request.RotationAngleDegrees
			};
		}

		/// <summary>
		/// Converts RedactionOptionsRequest to RedactionOptions
		/// </summary>
		/// <returns></returns>
		public static RedactionOptions GetOptions(this RedactionOptionsRequest request)
		{
			return new RedactionOptions
			{
				SearchQuery = request.SearchQuery,
				ReplaceText = request.ReplaceText,
				IsCaseSensitiveSearch = request.IsCaseSensitiveSearch,
				MustReplaceText = request.MustReplaceText,
				MustReplaceComments = request.MustReplaceComments,
				MustReplaceMetadata = request.MustReplaceMetadata
			};
		}

		public static PresentationMetadata GetModel(this PresentationMetadataDTO model)
		{
			return new PresentationMetadata
			{
				AppVersion = model.AppVersion,
				NameOfApplication = model.NameOfApplication,
				Company = model.Company,
				Manager = model.Manager,
				PresentationFormat = model.PresentationFormat,
				SharedDoc = model.SharedDoc,
				ApplicationTemplate = model.ApplicationTemplate,
				TotalEditingTime = model.TotalEditingTime ?? TimeSpan.Zero,
				Title = model.Title,
				Subject = model.Subject,
				Author = model.Author,
				Keywords = model.Keywords,
				Comments = model.Comments,
				Category = model.Category,
				CreatedTime = model.CreatedTime ?? DateTime.Now,
				LastSavedTime = model.LastSavedTime ?? DateTime.Now,
				LastPrinted = model.LastPrinted ?? DateTime.Now,
				LastSavedBy = model.LastSavedBy,
				RevisionNumber = model.RevisionNumber ?? 0,
				ContentStatus = model.ContentStatus,
				ContentType = model.ContentType,
				HyperlinkBase = model.HyperlinkBase,
				CustomProperties = ConvertCustomProperties(model.CustomProperties)
			};
		}

		public static object ToObject(JsonElement element)
		{
			if (element.ValueKind == JsonValueKind.String)
			{
				return element.ToObject<string>();
			}

			if (element.ValueKind == JsonValueKind.False || element.ValueKind == JsonValueKind.False)
			{
				return element.ToObject<bool>();
			}

			if (element.ValueKind == JsonValueKind.Number)
			{
				return element.ToObject<double>();
			}

			return null;
		}

		public static T ToObject<T>(this JsonElement element)
		{
			var json = element.GetRawText();
			return JsonSerializer.Deserialize<T>(json);
		}

		private static Dictionary<string, object> ConvertCustomProperties(Dictionary<string, object> properties)
		{
			var result = new Dictionary<string, object>();
			foreach (var name in properties.Keys)
			{
				object val = properties[name];
				result[name] = val is JsonElement element ? ToObject(element) : val;
			}
			return result;
		}
	}
}
