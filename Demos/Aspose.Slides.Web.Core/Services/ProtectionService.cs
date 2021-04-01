using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of slides protection logic.
	/// </summary>
	internal sealed class ProtectionService : SlidesServiceBase, IProtectionService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public ProtectionService(ILogger<ProtectionService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="password">Password.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public void Unlock(
			IList<string> sourceFiles,
			IList<string> outFiles,
			string password,
			CancellationToken cancellationToken = default
		)
		{
			void UnlockPresentationByIndex(int index)
			{
				var sourceFile = sourceFiles[index];
				var isViewProtected = false;

				try
				{
					using var testPresentation = new Presentation(sourceFile, new LoadOptions { OnlyLoadDocumentProperties = true });

					cancellationToken.ThrowIfCancellationRequested();

				}
				catch (InvalidPasswordException)
				{
					isViewProtected = true;
				}

				using var presentation = isViewProtected
						? new Presentation(sourceFile, new LoadOptions { Password = password })
						: new Presentation(sourceFile);

				cancellationToken.ThrowIfCancellationRequested();

				presentation.ProtectionManager.EncryptDocumentProperties = false;
				presentation.ProtectionManager.RemoveEncryption();
				presentation.ProtectionManager.RemoveWriteProtection();

				if (presentation.DocumentProperties.ContainsCustomProperty("_MarkAsFinal"))
					presentation.DocumentProperties.RemoveCustomProperty("_MarkAsFinal");

				cancellationToken.ThrowIfCancellationRequested();

				var outFile = outFiles[index];

				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile()); 
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, UnlockPresentationByIndex);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}
		}

		/// <summary>
		/// Asynchronously removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="password">Password.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task UnlockAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			string password,
			CancellationToken cancellationToken = default)
			=> await Task.Run(() => Unlock(sourceFiles, outFiles, password, cancellationToken));


		/// <summary>
		/// Applies protection to source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="markAsReadonly">Mark presentation as read-only.</param>
		/// <param name="markAsFinal">Mark presentation as final.</param>
		/// <param name="passwordEdit">Password for edit.</param>
		/// <param name="passwordView">Password for view.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public void Lock(
			IList<string> sourceFiles,
			IList<string> outFiles,
			bool markAsReadonly,
			bool markAsFinal,
			string passwordEdit,
			string passwordView,
			CancellationToken cancellationToken = default
		)
		{
			void LockPresentationByIndex(int index)
			{
				var sourceFile = sourceFiles[index];
				using var presentation = new Presentation(sourceFile);

				cancellationToken.ThrowIfCancellationRequested();

				if (!string.IsNullOrEmpty(passwordEdit))
					presentation.ProtectionManager.SetWriteProtection(passwordEdit);

				if (!string.IsNullOrEmpty(passwordView))
				{
					presentation.ProtectionManager.EncryptDocumentProperties = true;
					presentation.ProtectionManager.Encrypt(passwordView);
				}

				if (markAsReadonly)
					throw new NotImplementedException();

				if (markAsFinal)
					presentation.DocumentProperties.SetCustomPropertyValue("_MarkAsFinal", true);

				cancellationToken.ThrowIfCancellationRequested();

				var outFile = outFiles[index];

				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile()); 
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, LockPresentationByIndex);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}
		}

		/// <summary>
		/// Asynchronously applies protection to source file, saves resulted file to out file.
		/// If both view and edit passwords are blank, applies to file read-only.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="markAsReadonly">Mark presentation as read-only.</param>
		/// <param name="markAsFinal">Mark presentation as final.</param>
		/// <param name="passwordEdit">Password for edit.</param>
		/// <param name="passwordView">Password for view.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task LockAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			bool markAsReadonly,
			bool markAsFinal,
			string passwordEdit,
			string passwordView,
			CancellationToken cancellationToken = default
		) =>
			await Task.Run(() => Lock(
				sourceFiles,
				outFiles,
				markAsReadonly,
				markAsFinal,
				passwordEdit,
				passwordView,
				cancellationToken
			));
	}
}
