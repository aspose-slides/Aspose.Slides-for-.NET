using Aspose.Cells;
using Aspose.Slides.Charts;
using Aspose.Slides.Web.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal class BaseChartBuilder : IChartBuilder
	{
		public const string RowDefaultTitle = "Category";
		public const string ColumnDefaultTitle = "Series";

		protected const float DefaultChartStartPointX = 0f;
		protected const float DefaultChartStartPointY = 0f;
		protected const float DefaultChartWidth = 500f;
		protected const float DefaultChartHeight = 500f;
		
		protected readonly ChartType _chartType;

		private readonly List<KnownColor> _colorList;
		private readonly Random _rand;

		private int _maxColorIndex;
		private KnownColor _lastKnownColor;

		public BaseChartBuilder(ChartType chartType)
		{
			_chartType = chartType;

			_colorList = Enum.GetValues(typeof(KnownColor)).Cast<KnownColor>().ToList();
			_rand = new Random(DateTime.Now.Ticks.GetHashCode());
			_maxColorIndex = _colorList.Count();
		}

		public virtual void CreateChartForWorksheet(Worksheet worksheet, ISlide slide, MemoryStream memoryStream)
		{
			var lastRow = worksheet.Cells.Rows[worksheet.Cells.MaxDataRow];
			var lastCell = lastRow.GetCellOrNull(worksheet.Cells.MaxDataColumn);
			var formula = $"{worksheet.Name}!A1:{lastCell.Name}";
			var chart = GetChart(slide);

			chart.ChartData.Categories.Clear();
			chart.ChartData.Series.Clear();
			chart.ChartData.WriteWorkbookStream(memoryStream);
			chart.ChartData.SetRange(formula);
		}

		protected virtual IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height; 

			return slide.Shapes.AddChart(_chartType, DefaultChartStartPointX, DefaultChartStartPointY, chartWidth, chartHeight);
		}

		protected Color GetRandomColor()
		{
			KnownColor randomColorName;

			do
			{
				randomColorName = _colorList[_rand.Next(0, _maxColorIndex)];
			}
			while ( _lastKnownColor == randomColorName ||
			randomColorName.ToString().Contains("White") ||
			randomColorName.ToString().Contains("Transparent") ||
			randomColorName.ToString().Contains("Light"));

			_lastKnownColor = randomColorName;

			return Color.FromKnownColor(randomColorName);
		}
	}
}
