using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.Collections.Generic;
using System.IO;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class TreemapChartBuilder : BaseChartBuilder
	{
		private const float DefaultTreemapChartStartPointX = 0f;
		private const float DefaultTreemapChartStartPointY = 0f;
		private const float DefaultPreviewTreemapChartStartPointX = 0f;
		private const float DefaultPreviewTreemapChartStartPointY = 0f;

		public TreemapChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;
			var treemapChartStartPointX = chartWidth > DefaultChartWidth ? DefaultTreemapChartStartPointX : DefaultPreviewTreemapChartStartPointX;
			var treemapChartStartPointY = chartHeight > DefaultChartHeight ? DefaultTreemapChartStartPointY : DefaultPreviewTreemapChartStartPointY;

			return slide.Shapes.AddChart(_chartType, treemapChartStartPointX, treemapChartStartPointY, chartWidth, chartHeight);
		}

		public override void CreateChartForWorksheet(Worksheet worksheet, ISlide slide, MemoryStream memoryStream)
		{
			var chart = GetChart(slide);

			// Setting chart Title			
			chart.ChartTitle.AddTextFrameForOverriding(worksheet.Name);
			chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
			chart.ChartTitle.Height = 20;
			chart.HasTitle = true;

			// Delete default generated series and categories
			chart.ChartData.Series.Clear();
			chart.ChartData.Categories.Clear();

			int defaultWorksheetIndex = 0;

			// Getting the chart data worksheet
			IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

			fact.Clear(defaultWorksheetIndex);

			var rowCount = worksheet.Cells.MaxDataRow + 1;			
			var branches = new HashSet<string>();
			var stems = new HashSet<string>();
			var leaves = new HashSet<string>();

			var series = chart.ChartData.Series.Add(_chartType);
			series.Labels.DefaultDataLabelFormat.ShowCategoryName = true;
			series.ParentLabelLayout = ParentLabelLayoutType.Overlapping;

			for (var rowIndex = 1; rowIndex < rowCount; rowIndex++)
			{				
				var row = worksheet.Cells.Rows[rowIndex];
				IChartCategory leaf = null;								

				var cellBranch = row.GetCellOrNull(0);																									

				if (cellBranch == null || cellBranch.Type != CellValueType.IsString)
				{
					throw new ArgumentException($"The cell[{rowIndex}:0] has invalid input data.");
				}

				if (!branches.Contains(cellBranch.StringValue))
				{
					var cellStem = row.GetCellOrNull(1);

					if (cellStem == null || cellStem.Type != CellValueType.IsString)
					{
						throw new ArgumentException($"The cell[{rowIndex}:1] has invalid input data.");
					}

					if (!stems.Contains(cellStem.StringValue))
					{
						var cellLeaf = row.GetCellOrNull(2);

						if (cellLeaf == null || cellLeaf.Type != CellValueType.IsString)
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}

						if(!leaves.Contains(cellLeaf.StringValue))
						{
							leaf = chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, 2, cellLeaf.StringValue));

							var cellData = row.GetCellOrNull(3);

							if (cellData == null || cellData.Type != CellValueType.IsNumeric)
							{
								throw new ArgumentException($"The cell[{rowIndex}:3] has invalid input data.");
							}

							series.DataPoints.AddDataPointForTreemapSeries(fact.GetCell(defaultWorksheetIndex, rowIndex, 3, cellData.DoubleValue));
							leaves.Add(cellLeaf.StringValue);
						}
						else
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}

						leaf.GroupingLevels.SetGroupingItem(1, cellStem.StringValue);
						stems.Add(cellStem.StringValue);
					}

					leaf.GroupingLevels.SetGroupingItem(2, cellBranch.StringValue);
					branches.Add(cellBranch.StringValue);
				}
				else
				{
					var cellStem = row.GetCellOrNull(1);

					if (cellStem == null || cellStem.Type != CellValueType.IsString)
					{
						throw new ArgumentException($"The cell[{rowIndex}:1] has invalid input data.");
					}

					if (!stems.Contains(cellStem.StringValue))
					{
						var cellLeaf = row.GetCellOrNull(2);

						if (cellLeaf == null || cellLeaf.Type != CellValueType.IsString)
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}

						if (!leaves.Contains(cellLeaf.StringValue))
						{
							leaf = chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, 2, cellLeaf.StringValue));

							var cellData = row.GetCellOrNull(3);

							if (cellData == null || cellData.Type != CellValueType.IsNumeric)
							{
								throw new ArgumentException($"The cell[{rowIndex}:3] has invalid input data.");
							}

							series.DataPoints.AddDataPointForTreemapSeries(fact.GetCell(defaultWorksheetIndex, rowIndex, 3, cellData.DoubleValue));
							leaves.Add(cellLeaf.StringValue);
						}
						else
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}

						leaf.GroupingLevels.SetGroupingItem(1, cellStem.StringValue);
						stems.Add(cellStem.StringValue);
					}
					else
					{
						var cellLeaf = row.GetCellOrNull(2);

						if (cellLeaf == null || cellLeaf.Type != CellValueType.IsString)
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}

						if (!leaves.Contains(cellLeaf.StringValue))
						{
							chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, 2, cellLeaf.StringValue));

							var cellData = row.GetCellOrNull(3);

							if (cellData == null || cellData.Type != CellValueType.IsNumeric)
							{
								throw new ArgumentException($"The cell[{rowIndex}:3] has invalid input data.");
							}

							series.DataPoints.AddDataPointForTreemapSeries(fact.GetCell(defaultWorksheetIndex, rowIndex, 3, cellData.DoubleValue));
							leaves.Add(cellLeaf.StringValue);
						}
						else
						{
							throw new ArgumentException($"The cell[{rowIndex}:2] has invalid input data.");
						}
					}
				}
			}			
		}
	}
}
