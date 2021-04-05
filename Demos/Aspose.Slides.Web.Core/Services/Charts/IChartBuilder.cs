using Aspose.Cells;
using System.IO;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface of chart builder
	/// </summary>
	internal interface IChartBuilder
	{
		/// <summary>
		/// Creates Chart For Worksheet
		/// </summary>
		/// <param name="worksheet">Input data</param>
		/// <param name="slide">A slide of presentation</param>
		/// <param name="memoryStream">Input data stream</param>
		void CreateChartForWorksheet(Worksheet worksheet, ISlide slide, MemoryStream memoryStream);
	}
}
