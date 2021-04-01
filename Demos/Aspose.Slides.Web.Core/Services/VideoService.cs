using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services
{
	internal class VideoService : SlidesServiceBase, IVideoService
	{
		public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(3);

		public string FfmpegPath { get; }

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="ffmpegPath"></param>
		/// <param name="licenseProvider"></param>
		internal VideoService(ILogger<VideoService> logger,  string ffmpegPath, ILicenseProvider licenseProvider) : base(logger)
		{
			FfmpegPath = ffmpegPath ?? throw new ArgumentException(nameof(ffmpegPath));
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		public string Encode(
			string sourceFile,
			string outFolder,
			string splitRange,
			int transitionTime,
			VideoCodecs codec,
			CancellationToken cancellationToken = default)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				cancellationToken.ThrowIfCancellationRequested();

				IEnumerable<(string name, ISlide[] slides)> chunks = string.IsNullOrEmpty(splitRange) ?
					presentation.GetChunks(SplitTypes.SlideBySlide, 0, null) :
					presentation.GetChunks(SplitTypes.Range, 0, splitRange);

				var frameCounter = 0;

				foreach (var chunk in chunks)
				{
					foreach (var slide in chunk.slides)
					{
						cancellationToken.ThrowIfCancellationRequested();

						var outFile = Path.Combine(outFolder, $"img{frameCounter:D3}.png");
						using (var bitmap = slide.GetThumbnail(1, 1))
						{
							bitmap.Save(outFile, ImageFormat.Png);
						}

						frameCounter++;
					}
				}
			}

			var result = Path.Combine(outFolder, $"{Path.GetFileNameWithoutExtension(sourceFile)}.mp4");

			cancellationToken.ThrowIfCancellationRequested();

			EncodeVideo(outFolder, result, codec, transitionTime);

			return result;
		}

		private void EncodeVideo(string outFolder, string result, VideoCodecs codec, int transitionTime)
		{
			string codecName = codec == VideoCodecs.H265 ? "libx265" : "libx264";
			using var ffmpeg = new Process
			{
				StartInfo =
				{
					UseShellExecute = false,
					FileName = FfmpegPath,
					Arguments =
						$"-r 1/{transitionTime} -i \"{outFolder}\\img%03d.png\" -c:v {codecName} -vf fps=25 -pix_fmt yuv420p \"{result}\""
				}
			};
			try
			{
				ffmpeg.Start();
			}
			catch (Win32Exception e)
			{
				if (e.NativeErrorCode == 2)
				{
					throw new VideoConversionException(
						"FFmpeg.exe is missing. It can be downloaded from here: https://www.gyan.dev/ffmpeg/builds/"
					);
				}

				throw;
			}

			var exited = ffmpeg.WaitForExit((int)Timeout.TotalMilliseconds);
			if (exited)
			{
				if (ffmpeg.ExitCode != 0)
				{
					throw new VideoConversionException($"FFmpeg exited with the code {ffmpeg.ExitCode}.");
				}
			}
			else
			{
				ffmpeg.Kill();
				ffmpeg.WaitForExit();
				throw new ProcessingTimeoutException("Video conversion timeout");
			}
		}
	}
}
