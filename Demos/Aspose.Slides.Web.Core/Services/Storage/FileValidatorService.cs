using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	/// <summary>
	/// The implementation of validation logic
	/// </summary>
	public sealed class FileValidatorService : SlidesServiceBase, IFileValidatorService
	{
		private readonly Dictionary<byte[], Tuple<int, string>> _magicNumbersToFileTypes;
		private readonly IEnumerable<string> _validFileTypes;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="validFileTypes">Valid file types for the input file. For ex.: exe, pptx, elf or !#(scripts)</param>
		public FileValidatorService(ILogger<FileValidatorService> logger, IEnumerable<string> validFileTypes) : base(logger)
		{
			if (validFileTypes == null || !validFileTypes.Any())
			{
				throw new ArgumentNullException(nameof(validFileTypes));
			}

			_validFileTypes = validFileTypes;

			_magicNumbersToFileTypes = new Dictionary<byte[], Tuple<int, string>>();

			_magicNumbersToFileTypes.Add(new byte[] { 0xED, 0xAB, 0xEE, 0xDB }, new Tuple<int, string>(0,"rpm"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x53, 0x50, 0x30, 0x31 }, new Tuple<int, string>(0,"bin"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00 }, new Tuple<int, string>(0, "pic, pif, sea, ytr"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
														0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
														0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, new Tuple<int, string>(11, "pdb"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xBE, 0xBA, 0xFE, 0xCA }, new Tuple<int, string>(0, "dba"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x01, 0x42, 0x44 }, new Tuple<int, string>(0,"dba"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x01, 0x44, 0x54 }, new Tuple<int, string>(0, "tba"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x00, 0x01, 0x11 }, new Tuple<int, string>(0,"ico"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x66, 0x74, 0x79, 0x70, 0x33, 0x67 }, new Tuple<int, string>(4, "3gp, 3g2"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x1F, 0x9D }, new Tuple<int, string>(0, "z, tar.z"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x1F, 0xA0 }, new Tuple<int, string>(0, "z, tar.z"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x42, 0x41, 0x43, 0x4B, 0x4D, 0x49, 0x4B, 0x45, 0x44, 0x49, 0x53, 0x4B }, new Tuple<int, string>(0, "bac"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x42, 0x5A, 0x68 }, new Tuple<int, string>(0, "bz2"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x37, 0x61 }, new Tuple<int, string>(0, "gif"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, new Tuple<int, string>(0, "gif"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x49, 0x49, 0x2A, 0x00 }, new Tuple<int, string>(0, "tiff"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x4D, 0x00, 0x2A }, new Tuple<int, string>(0, "tiff"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x10, 0x00, 0x00, 0x00, 0x43, 0x52 }, new Tuple<int, string>(0, "cr2"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x52, 0x4E, 0x43, 0x01 }, new Tuple<int, string>(0, "rnc"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x52, 0x4E, 0x43, 0x02 }, new Tuple<int, string>(0, "rnc"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x53, 0x44, 0x50, 0x58 }, new Tuple<int, string>(0, "dpx"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x58, 0x50, 0x44, 0x53 }, new Tuple<int, string>(0, "dpx"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x76, 0x2F, 0x31, 0x01 }, new Tuple<int, string>(0, "exr"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x42, 0x50, 0x47, 0xFB }, new Tuple<int, string>(0, "bpg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xD8, 0xFF, 0xDB }, new Tuple<int, string>(0, "jpg, jpeg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }, new Tuple<int, string>(0, "jpg, jpeg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xD8, 0xFF, 0xE1 }, new Tuple<int, string>(0, "jpg, jpeg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x49, 0x4E, 0x44, 0x58 }, new Tuple<int, string>(0, "idx"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4C, 0x5A, 0x49, 0x50 }, new Tuple<int, string>(0, "lz"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x5A }, new Tuple<int, string>(0, "exe, dll"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, new Tuple<int, string>(0, "zip, jar, odt, ods, odp, docx, xlsx, pptx, vsdx,apk, aar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x50, 0x4B, 0x05, 0x06 }, new Tuple<int, string>(0, "zip, jar, odt, ods, odp, docx, xlsx, pptx, vsdx,apk, aar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x50, 0x4B, 0x07, 0x08 }, new Tuple<int, string>(0, "zip, jar, odt, ods, odp, docx, xlsx, pptx, vsdx,apk, aar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x52, 0x61, 0x72, 0x21, 0x1A, 0x07, 0x00 }, new Tuple<int, string>(0, "rar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x52, 0x61, 0x72, 0x21, 0x1A, 0x07, 0x01, 0x00 }, new Tuple<int, string>(0, "rar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x7F, 0x45, 0x4C, 0x46 }, new Tuple<int, string>(0, "elf"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, new Tuple<int, string>(0, "png"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xCA, 0xFE, 0xBA, 0xBE }, new Tuple<int, string>(0, "class"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x25, 0x21, 0x50, 0x53 }, new Tuple<int, string>(0, "ps"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x25, 0x50, 0x44, 0x46 }, new Tuple<int, string>(0, "pdf"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x30, 0x26, 0xB2, 0x75, 0x8E, 0x66, 0xCF, 0x11,
														0xA6, 0xD9, 0x00, 0xAA, 0x00, 0x62, 0xCE, 0x6C}, new Tuple<int, string>(0, "asf, wma, wmv"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4F, 0x67, 0x67, 0x53 }, new Tuple<int, string>(0, "ogg, oga, ogv"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x38, 0x42, 0x50, 0x53 }, new Tuple<int, string>(0, "psd"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x52, 0x49, 0x46, 0x46 }, new Tuple<int, string>(0, "wav, avi"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xFB }, new Tuple<int, string>(0, "mp3"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xF3 }, new Tuple<int, string>(0, "mp3"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xFF, 0xF2 }, new Tuple<int, string>(0, "mp3"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x49, 0x44, 0x33 }, new Tuple<int, string>(0, "mp3"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x42, 0x4D }, new Tuple<int, string>(0, "bmp, dib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x43, 0x44, 0x30, 0x30, 0x31 }, new Tuple<int, string>(0, "iso"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x66, 0x4C, 0x61, 0x43 }, new Tuple<int, string>(0, "flac"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x54, 0x68, 0x64 }, new Tuple<int, string>(0, "mid, midi"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, new Tuple<int, string>(0, "doc, xls, ppt, msg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x64, 0x65, 0x78, 0x0A, 0x30, 0x33, 0x35, 0x00 }, new Tuple<int, string>(0, "dex"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4B, 0x44, 0x4D }, new Tuple<int, string>(0, "vmdk"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x43, 0x72, 0x32, 0x34 }, new Tuple<int, string>(0, "crx"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x41, 0x47, 0x44, 0x33 }, new Tuple<int, string>(0, "fh8"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x05, 0x07, 0x00, 0x00, 0x42, 0x4F, 0x42, 0x4F,
														0x05, 0x07, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
														0x00, 0x00, 0x00,0x00, 0x00, 0x01}, new Tuple<int, string>(0, "cwk"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x06, 0x07, 0xE1, 0x00, 0x42, 0x4F, 0x42, 0x4F,
														0x06, 0x07, 0xE1, 0x00, 0x00, 0x00, 0x00, 0x00,
														0x00, 0x00, 0x00,0x00, 0x00, 0x01}, new Tuple<int, string>(0, "cwk"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x45, 0x52, 0x02, 0x00, 0x00, 0x00 }, new Tuple<int, string>(0, "toast"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x8B, 0x45, 0x52, 0x02, 0x00, 0x00, 0x00 }, new Tuple<int, string>(0, "toast"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x01, 0x73, 0x0D, 0x62, 0x62, 0x60 }, new Tuple<int, string>(0, "dmg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x61, 0x72, 0x21 }, new Tuple<int, string>(0, "xar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x50, 0x4D, 0x4F, 0x43, 0x43, 0x4D, 0x4F, 0x43 }, new Tuple<int, string>(0, "dat"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4E, 0x45, 0x53, 0x1A }, new Tuple<int, string>(0, "nes"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x75, 0x73, 0x74, 0x61, 0x72, 0x00, 0x30, 0x30 }, new Tuple<int, string>(0, "tar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x75, 0x73, 0x74, 0x61, 0x72, 0x20, 0x20, 0x00 }, new Tuple<int, string>(0, "tar"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x74, 0x6F, 0x78, 0x33 }, new Tuple<int, string>(0, "tox"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x4C, 0x56, 0x49 }, new Tuple<int, string>(0, "mlv"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C }, new Tuple<int, string>(0, "7z"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x1F, 0x8B }, new Tuple<int, string>(0, "gz, tar.gz"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x04, 0x22, 0x4D, 0x18 }, new Tuple<int, string>(0, "lz4"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x53, 0x43, 0x46 }, new Tuple<int, string>(0, "cab"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x46, 0x4C, 0x49, 0x46 }, new Tuple<int, string>(0, "flif"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x1A, 0x45, 0xDF, 0xA3 }, new Tuple<int, string>(0, "mkv, mka, mks, mk3d, webm"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4D, 0x49, 0x4C, 0x20 }, new Tuple<int, string>(0, "stg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x41, 0x54, 0x26, 0x54, 0x46, 0x4F, 0x52, 0x4D }, new Tuple<int, string>(0, "djvu, djv"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x30, 0x82 }, new Tuple<int, string>(0, "der"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x44, 0x49, 0x43, 0x4D }, new Tuple<int, string>(0, "dcm"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x77, 0x4F, 0x46, 0x46 }, new Tuple<int, string>(0, "woff"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x77, 0x4F, 0x46, 0x32 }, new Tuple<int, string>(0, "woff2"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x3C, 0x3F, 0x78, 0x6D, 0x6C, 0x20 }, new Tuple<int, string>(0, "xml"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x6D, 0x73, 0x61, 0x00 }, new Tuple<int, string>(0, "wasm"));
			_magicNumbersToFileTypes.Add(new byte[] { 0xCF, 0x84, 0x01 }, new Tuple<int, string>(0, "lep"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x43, 0x57, 0x53 }, new Tuple<int, string>(0, "swf"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x46, 0x57, 0x53 }, new Tuple<int, string>(0, "swf"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x21, 0x3C, 0x61, 0x72, 0x63, 0x68, 0x3E }, new Tuple<int, string>(0, "deb"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x7B, 0x5C, 0x72, 0x74, 0x66, 0x31 }, new Tuple<int, string>(0, "rtf"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x47 }, new Tuple<int, string>(0, "ts, tsv, tsa"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x00, 0x01, 0xBA }, new Tuple<int, string>(0, "m2p, vob"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x00, 0x00, 0x01, 0xB3 }, new Tuple<int, string>(0, "mpg, mpeg"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x01 }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x5E }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x9C }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0xDA }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x20 }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0x7D }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0xBB }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x78, 0xF9 }, new Tuple<int, string>(0, "zlib"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x1F, 0x8B, 0x08, 0x00 }, new Tuple<int, string>(1, "dat"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x63, 0x76, 0x78, 0x32 }, new Tuple<int, string>(0, "lzfse"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4F, 0x52, 0x43 }, new Tuple<int, string>(0, "orc"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x4F, 0x62, 0x6A, 0x01 }, new Tuple<int, string>(0, "avro"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x53, 0x45, 0x51, 0x36 }, new Tuple<int, string>(0, "rc"));
			_magicNumbersToFileTypes.Add(new byte[] { 0x23, 0x21 }, new Tuple<int, string>(0, "#!"));
		}

		/// <summary>
		/// Validates files
		/// </summary>
		/// <param name="fileName">The input file for validation</param>		
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Is valid or not</returns>
		public bool IsValidFile(string fileName, CancellationToken cancellationToken = default)
		{
			if(String.IsNullOrWhiteSpace(fileName))
			{
				throw new ArgumentNullException(nameof(fileName));
			}

			if(!File.Exists(fileName))
			{
				throw new ArgumentNullException(nameof(fileName));
			}				

			var validFileTypes = _validFileTypes.Select(ve => ve.ToLower());

			cancellationToken.ThrowIfCancellationRequested();

			return IsValidByFileExtensions(fileName, validFileTypes) &&				
				IsValidateFileTypeByMagicNumber(fileName, validFileTypes);			
		}

		private bool IsValidByFileExtensions(string fileName, IEnumerable<string> validFileTypes)
		{
			var fileExtension = Path.GetExtension(fileName).TrimStart(new char[] {'.'}).ToLowerInvariant();

			return String.IsNullOrWhiteSpace(fileExtension) ? true : validFileTypes.Contains(fileExtension);			
		}		

		// See the https://en.wikipedia.org/wiki/List_of_file_signatures
		private bool IsValidateFileTypeByMagicNumber(string fileName, IEnumerable<string> validFileTypes, CancellationToken cancellationToken = default)
		{						
			var fileBytes = File.ReadAllBytes(fileName);

			foreach (var magicNumberToFileType in _magicNumbersToFileTypes)
			{
				cancellationToken.ThrowIfCancellationRequested();
				
				var testBytes = new byte[magicNumberToFileType.Key.Length];
				var offset = magicNumberToFileType.Value.Item1;

				if (fileBytes.Length < offset + testBytes.Length)
				{
					continue;
				}

				Array.Copy(fileBytes, offset, testBytes, 0, magicNumberToFileType.Key.Length);

				var isSameMagicNumber = magicNumberToFileType.Key.Compare(testBytes);

				if(!isSameMagicNumber)
				{
					continue;
				}

				var types = magicNumberToFileType.Value.Item2.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim().ToLower());
				var findTypes = validFileTypes.Intersect(types);

				return findTypes.Any() ? true : false;
			}

			return true;
		}

		/// <summary>
		/// Validates files asynchronously
		/// </summary>
		/// <param name="fileName">The input file for validation</param>		
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Is valid or not</returns>
		public async Task<bool> IsValidFileAsync(string fileName, CancellationToken cancellationToken = default)
			=> await Task.Run(() => IsValidFile(fileName, cancellationToken));
	}
}
