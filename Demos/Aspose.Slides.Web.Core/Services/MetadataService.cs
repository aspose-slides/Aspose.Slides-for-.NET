using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Models.Metadata;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of metadata logic.
	/// </summary>
	internal sealed class MetadataService : SlidesServiceBase, IMetadataService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public MetadataService(ILogger<MetadataService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Gets presentation metadata.
		/// </summary>
		/// <param name="sourceFile">Path to the presentation file.</param>
		/// <returns>Metadata object.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public PresentationMetadata GetMetadata(string sourceFile, CancellationToken cancellationToken = default)
		{
			using var presentation = new Presentation(sourceFile);
			
			cancellationToken.ThrowIfCancellationRequested();

			return CreatePresentationMetadata(presentation.DocumentProperties, cancellationToken);			
		}

		/// <summary>
		/// Updates presentation metadata.
		/// </summary>
		/// <param name="sourceFile">Path to the source presentation file.</param>
		/// <param name="outFile">Path to the resulting presentation file with applied metadata.</param>
		/// <param name="metadata">Metadata object.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public void UpdateMetadata(
			string sourceFile,
			string outFile,
			PresentationMetadata metadata,
			CancellationToken cancellationToken = default
		)
		{
			if (metadata == null)
			{
				throw new ArgumentException("Metadata object is null.");
			}

			using var presentation = new Presentation(sourceFile);
			
			cancellationToken.ThrowIfCancellationRequested();

			presentation.DocumentProperties.ApplicationTemplate = metadata.ApplicationTemplate;
			presentation.DocumentProperties.Author = metadata.Author;
			presentation.DocumentProperties.Category = metadata.Category;
			presentation.DocumentProperties.Comments = metadata.Comments;
			presentation.DocumentProperties.Company = metadata.Company;
			presentation.DocumentProperties.ContentStatus = metadata.ContentStatus;
			presentation.DocumentProperties.ContentType = metadata.ContentType;
			presentation.DocumentProperties.HyperlinkBase = metadata.HyperlinkBase;
			presentation.DocumentProperties.Keywords = metadata.Keywords;
			presentation.DocumentProperties.LastSavedBy = metadata.LastSavedBy;
			presentation.DocumentProperties.Manager = metadata.Manager;
			presentation.DocumentProperties.NameOfApplication = metadata.NameOfApplication;
			presentation.DocumentProperties.PresentationFormat = metadata.PresentationFormat;
			presentation.DocumentProperties.Subject = metadata.Subject;
			presentation.DocumentProperties.Title = metadata.Title;
			presentation.DocumentProperties.CreatedTime = metadata.CreatedTime;
			presentation.DocumentProperties.LastPrinted = metadata.LastPrinted;
			presentation.DocumentProperties.RevisionNumber = metadata.RevisionNumber;
			presentation.DocumentProperties.SharedDoc = metadata.SharedDoc;
			presentation.DocumentProperties.TotalEditingTime = metadata.TotalEditingTime;
			presentation.DocumentProperties.ClearCustomProperties();

			if (metadata.CustomProperties != null)
			{
				foreach (var name in metadata.CustomProperties.Keys)
				{
					if (name != null && metadata.CustomProperties[name] != null)
					{
						presentation.DocumentProperties[name] = metadata.CustomProperties[name] is long longVal ? (int) longVal : metadata.CustomProperties[name];
					}

					cancellationToken.ThrowIfCancellationRequested();
				}
			}

			cancellationToken.ThrowIfCancellationRequested();

			presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());
		}

		private PresentationMetadata CreatePresentationMetadata(IDocumentProperties properties, CancellationToken cancellationToken = default)
		{
			var metadata = new PresentationMetadata
			{
				AppVersion = properties.AppVersion,
				ApplicationTemplate = properties.ApplicationTemplate,
				Author = properties.Author,
				Category = properties.Category,
				Comments = properties.Comments,
				Company = properties.Company,
				ContentStatus = properties.ContentStatus,
				ContentType = properties.ContentType,
				HyperlinkBase = properties.HyperlinkBase,
				Keywords = properties.Keywords,
				LastSavedBy = properties.LastSavedBy,
				Manager = properties.Manager,
				NameOfApplication = properties.NameOfApplication,
				PresentationFormat = properties.PresentationFormat,
				Subject = properties.Subject,
				Title = properties.Title,
				CreatedTime = properties.CreatedTime,
				LastPrinted = properties.LastPrinted,
				LastSavedTime = properties.LastSavedTime,
				RevisionNumber = properties.RevisionNumber,
				SharedDoc = properties.SharedDoc,
				TotalEditingTime = properties.TotalEditingTime,
			};

			var customProperties = new Dictionary<string, object>();
			for (int i = 0; i < properties.CountOfCustomProperties; ++i)
			{
				string name = properties.GetCustomPropertyName(i);
				customProperties[name] = properties[name];

				cancellationToken.ThrowIfCancellationRequested();
			}

			metadata.CustomProperties = customProperties;

			return metadata;
		}
	}
}
