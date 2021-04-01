using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Microsoft.Extensions.Logging;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Conversion
{
	/// <summary>
	/// Implementation of gif encoder logic.
	/// All used classes available on .NET 5
	/// </summary>
	internal class GifEncoder : IGifEncoder
	{
		private readonly ILogger _logger;
		private readonly byte[] buf2 = new byte[19];
		private readonly byte[] buf3 = new byte[8];

		/// <summary>
		/// Ctor
		/// </summary>
		public GifEncoder(ILogger<GifEncoder> logger, ILicenseProvider licenseProvider)
		{
			_logger = logger;
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);

			buf2[0] = 33;  //extension introducer
			buf2[1] = 255; //application extension
			buf2[2] = 11;  //size of block
			buf2[3] = 78;  //N
			buf2[4] = 69;  //E
			buf2[5] = 84;  //T
			buf2[6] = 83;  //S
			buf2[7] = 67;  //C
			buf2[8] = 65;  //A
			buf2[9] = 80;  //P
			buf2[10] = 69; //E
			buf2[11] = 50; //2
			buf2[12] = 46; //.
			buf2[13] = 48; //0
			buf2[14] = 3;  //Size of block
			buf2[15] = 1;  //
			buf2[16] = 0;  //
			buf2[17] = 0;  //
			buf2[18] = 0;  //Block terminator
			buf3[0] = 33;  //Extension introducer
			buf3[1] = 249; //Graphic control extension
			buf3[2] = 4;   //Size of block
			buf3[3] = 9;   //Flags: reserved, disposal method, user input, transparent color
			buf3[4] = 1;  //Delay time low byte
			buf3[5] = 0;   //Delay time high byte
			buf3[6] = 255; //Transparent color index
			buf3[7] = 0;   //Block terminator
		}

		/// <summary>
		/// Encodes to gif format.
		/// </summary>
		/// <param name="presentation"></param>
		/// <param name="outputFileName"></param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns></returns>
		public string Encode(IPresentation presentation, string outputFileName, CancellationToken cancellationToken = default)
		{
			// See https://www.rickvandenbosch.net/blog/howto-create-an-animated-gif-using-net-c/
			byte[] buf1;
			using var memoryStream = new MemoryStream();
			using BinaryWriter binaryWriter = new BinaryWriter(new FileStream(outputFileName, FileMode.Create));

			for (var slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++)
			{
				cancellationToken.ThrowIfCancellationRequested();

				var slide = presentation.Slides[slideIndex];
				using var bitMap = slide.GetThumbnail(1, 1);

				bitMap.Save(memoryStream, ImageFormat.Gif);
				buf1 = memoryStream.ToArray();

				if (slideIndex == 0)
				{
					//only write these the first time….
					binaryWriter.Write(buf1, 0, 781); //Header & global color table
					binaryWriter.Write(buf2, 0, 19); //Application extension
				}

				binaryWriter.Write(buf3, 0, 8); //Graphic extension
				binaryWriter.Write(buf1, 789, buf1.Length - 790); //Image data

				if (slideIndex == presentation.Slides.Count - 1)
				{
					//only write this one the last time….
					binaryWriter.Write(";"); //Image terminator
				}

				memoryStream.SetLength(0);
			}

			binaryWriter.Close();
			memoryStream.Close();

			return outputFileName;
		}

		/// <summary>
		/// Encodes to gif format asynchronously.
		/// </summary>
		/// <param name="presentation"></param>
		/// <param name="outputFileName"></param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns></returns>
		public async Task<string> EncodeAsync(IPresentation presentation, string outputFileName, CancellationToken cancellationToken = default)
			=>  await Task.Run(() => Encode(presentation, outputFileName, cancellationToken));
	}
}
