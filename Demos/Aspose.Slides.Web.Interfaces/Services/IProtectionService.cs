using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides protection logic.
	/// </summary>
	public interface IProtectionService
	{
		/// <summary>
		/// Removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="password">Password.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		void Unlock(
			IList<string> sourceFiles,
			IList<string> outFiles,
			string password,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously removes password protection from source file, saves resulted file to out file.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="password">Password.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task UnlockAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			string password,
			CancellationToken cancellationToken = default);

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
		void Lock(
			IList<string> sourceFiles,
			IList<string> outFiles,
			bool markAsReadonly,
			bool markAsFinal,
			string passwordEdit,
			string passwordView,
			CancellationToken cancellationToken = default
		);

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
		Task LockAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			bool markAsReadonly,
			bool markAsFinal,
			string passwordEdit,
			string passwordView,
			CancellationToken cancellationToken = default
		);
	}
}
