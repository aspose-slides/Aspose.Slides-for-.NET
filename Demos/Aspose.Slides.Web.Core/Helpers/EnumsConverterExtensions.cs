using System;
using Aspose.Slides.Charts;
using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.Core.Helpers
{
	public static class EnumsConverterExtensions
	{
		private const string Jpg = "jpg";

		public static ChartType ConvertEnum(this ChartTypes chartType)
		{
			return chartType.ToString().ParseEnum<ChartType>(true);
		}
		
		public static SlidesConversionFormats ToSlidesConversionFormats(this string source)
		{
			source = source.TrimStart('.').ToLowerInvariant();

			if (source.Contains(Jpg, StringComparison.InvariantCultureIgnoreCase))
			{
				return SlidesConversionFormats.jpeg;
			}

			return source.ParseEnum<SlidesConversionFormats>();
		}
	}
}
