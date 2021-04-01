using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	internal sealed class PresentationTemplateService : SlidesServiceBase, IPresentationTemplateService
	{
		private readonly HttpClient _httpClient;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="basePath"></param>
		internal PresentationTemplateService(ILogger<PresentationTemplateService> logger, string basePath) : base(logger)
		{
			if(string.IsNullOrWhiteSpace(basePath))
			{
				throw new ArgumentNullException(nameof(basePath));
			}

			// https://stackoverflow.com/questions/23438416/why-is-httpclient-baseaddress-not-working
			if(!basePath.EndsWith('/'))
			{
				basePath += "/"; 
			}

			_httpClient = new HttpClient {BaseAddress = new Uri(basePath)};
		}

		public async Task<Stream> GetTemplateStreamAsync(string template, CancellationToken cancellationToken)
		{
			using var response = await _httpClient.GetAsync(template, cancellationToken);

			if (!response.IsSuccessStatusCode)
			{
				response.Dispose();

				return null;
			}

			// memStream will be disposed outside
			using var stream = await response.Content.ReadAsStreamAsync();
			var memStream = new MemoryStream((int)stream.Length);

			await stream.CopyToAsync(memStream, cancellationToken);
			
			return memStream;
		}
	}
}
