using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Web;


namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="password">Password.</param>
		public void Unlock(
			string sourceFile,
			string outFile,
			string password
		)
		{
			var isViewProtected = false;

			try
			{
				using (var presentation = new Presentation(sourceFile, new LoadOptions { OnlyLoadDocumentProperties = true }))
				{
				}
			}
			catch (InvalidPasswordException)
			{
				isViewProtected = true;
			}

			using (
				var presentation =
					isViewProtected
					? new Presentation(sourceFile, new LoadOptions { Password = password })
					: new Presentation(sourceFile)
			)
			{
				presentation.ProtectionManager.EncryptDocumentProperties = false;
				presentation.ProtectionManager.RemoveEncryption();
				presentation.ProtectionManager.RemoveWriteProtection();

				if (presentation.DocumentProperties.ContainsCustomProperty("_MarkAsFinal"))
					presentation.DocumentProperties.RemoveCustomProperty("_MarkAsFinal");

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}
		}

		/// <summary>
		/// Asynchronously removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="password">Password.</param>
		public  Task UnlockFile(string sourceFile, string outFile, string password)
			=>  Task.Run(() => Unlock(sourceFile, outFile, password));


		/// <summary>
		/// Applies protection to source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="markAsReadonly">Mark presentation as read-only.</param>
		/// <param name="markAsFinal">Mark presentation as final.</param>
		/// <param name="passwordEdit">Password for edit.</param>
		/// <param name="passwordView">Password for view.</param>
		public void Lock(
			string sourceFile,
			string outFile,
			bool markAsReadonly,
			bool markAsFinal,
			string passwordEdit,
			string passwordView
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
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

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}
		}

		/// <summary>
		/// Asynchronously applies protection to source file, saves resulted file to out file.
		/// If both view and edit passwords are blank, applies to file read-only.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="markAsReadonly">Mark presentation as read-only.</param>
		/// <param name="markAsFinal">Mark presentation as final.</param>
		/// <param name="passwordEdit">Password for edit.</param>
		/// <param name="passwordView">Password for view.</param>
		public  Task LockFile(string sourceFile, string outFile, bool markAsReadonly, bool markAsFinal, string passwordEdit, string passwordView)
			=>  Task.Run(() => Lock(
				sourceFile: sourceFile,
				outFile: outFile,
				markAsReadonly: markAsReadonly,
				markAsFinal: markAsFinal,
				passwordEdit: passwordEdit,
				passwordView: passwordView
			));
	}
}
