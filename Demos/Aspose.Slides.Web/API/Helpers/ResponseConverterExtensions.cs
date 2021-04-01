using Aspose.Slides.Web.API.Clients.DTO;
using Aspose.Slides.Web.Interfaces.Models.Metadata;
using System.Collections.Generic;

namespace Aspose.Slides.Web.API.Helpers
{
	internal static class ResponseConverterExtensions
	{
		public static PresentationMetadataDTO GetDTO(this PresentationMetadata model)
		{
			return new PresentationMetadataDTO
			{
				AppVersion = model.AppVersion,
				NameOfApplication = model.NameOfApplication,
				Company = model.Company,
				Manager = model.Manager,
				PresentationFormat = model.PresentationFormat,
				SharedDoc = model.SharedDoc,
				ApplicationTemplate = model.ApplicationTemplate,
				TotalEditingTime = model.TotalEditingTime,
				Title = model.Title,
				Subject = model.Subject,
				Author = model.Author,
				Keywords = model.Keywords,
				Comments = model.Comments,
				Category = model.Category,
				CreatedTime = model.CreatedTime,
				LastSavedTime = model.LastSavedTime,
				LastPrinted = model.LastPrinted,
				LastSavedBy = model.LastSavedBy,
				RevisionNumber = model.RevisionNumber,
				ContentStatus = model.ContentStatus,
				ContentType = model.ContentType,
				HyperlinkBase = model.HyperlinkBase,
				CustomProperties = new Dictionary<string, object>(model.CustomProperties)
			};
		}

		public static PresentationInfoDTO GetDTO(this Aspose.Slides.Web.Interfaces.Models.Viewer.PresentationInfo model)
		{
			return new PresentationInfoDTO
			{
				Width = model.Width,
				Height  = model.Height,
				Count = model.Count
			};
		}
	}
}
