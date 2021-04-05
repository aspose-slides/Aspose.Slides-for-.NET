using System.ComponentModel;

namespace Aspose.Slides.Web.API.Clients.Enums
{
	/// <summary>
	///  Represents a type of chart.
	/// </summary>
	public enum ChartTypes
	{
		/// <summary>
		/// Represents Clustered Column Chart.
		/// </summary>
		[Description("Clustered column")]
		ClusteredColumn = 0,

		/// <summary>
		/// Represents Stacked Column Chart.
		/// </summary>
		[Description("Stacked column")]
		StackedColumn = 1,

		/// <summary>
		/// Represents 100% Stacked Column Chart.
		/// </summary>
		[Description("Percents stacked column")]
		PercentsStackedColumn = 2,

		/// <summary>
		/// Represents 3D Colustered Column Chart.
		/// </summary>
		[Description("Clustered 3D column")]
		ClusteredColumn3D = 3,

		/// <summary>
		/// Represents 3D Stacked Column Chart.
		/// </summary>
		[Description("Stacked 3D column")]
		StackedColumn3D = 4,

		/// <summary>
		/// Represents 3D 100% Stacked Column Chart.
		/// </summary>
		[Description("Percents stacked 3D column")]
		PercentsStackedColumn3D = 5,

		/// <summary>
		/// Represents 3D Column Chart.
		/// </summary>
		[Description("Column 3D")]
		Column3D = 6,

		/// <summary>
		/// Represents Cylinder Chart.
		/// </summary>
		[Description("Clustered cylinder")]
		ClusteredCylinder = 7,

		/// <summary>
		/// Represents Stacked Cylinder Chart.
		/// </summary>
		[Description("Stacked cylinder")]
		StackedCylinder = 8,

		/// <summary>
		/// Represents 100% Stacked Cylinder Chart.
		/// </summary>
		[Description("Percents stacked cylinder")]
		PercentsStackedCylinder = 9,

		/// <summary>
		/// Represents 3D Cylindrical Column Chart.
		/// </summary>
		[Description("Cylinder 3D")]
		Cylinder3D = 10,

		/// <summary>
		/// Represents Cone Chart.
		/// </summary>
		[Description("Clustered cone")]
		ClusteredCone = 11,

		/// <summary>
		/// Represents Stacked Cone Chart.
		/// </summary>
		[Description("Stacked cone")]
		StackedCone = 12,

		/// <summary>
		/// Represents 100% Stacked Cone Chart.
		/// </summary>
		[Description("Percents stacked cone")]
		PercentsStackedCone = 13,

		/// <summary>
		/// Represents 3D Conical Column Chart.
		/// </summary>
		[Description("Cone 3D")]
		Cone3D = 14,

		/// <summary>
		/// Represents Pyramid Chart.
		/// </summary>
		[Description("Clustered pyramid")]
		ClusteredPyramid = 15,

		/// <summary>
		/// 
		/// </summary>
		[Description("Stacked pyramid")]
		StackedPyramid = 16,

		/// <summary>
		/// Represents 100% Stacked Pyramid Chart.
		/// </summary>
		[Description("Percents stacked pyramid")]
		PercentsStackedPyramid = 17,

		/// <summary>
		/// Represents 3D Pyramid Column Chart.
		/// </summary>
		[Description("Pyramid 3D")]
		Pyramid3D = 18,

		/// <summary>
		/// Represents Line Chart.
		/// </summary>
		[Description("Line")]
		Line = 19,

		/// <summary>
		/// Represents Stacked Line Chart.
		/// </summary>
		[Description("Stacked Line")]
		StackedLine = 20,

		/// <summary>
		/// Represents 100% Stacked Line Chart.
		/// </summary>
		[Description("Percents stacked line")]
		PercentsStackedLine = 21,

		/// <summary>
		/// Represents Line Chart with data markers.
		/// </summary>
		[Description("Line with markers")]
		LineWithMarkers = 22,

		/// <summary>
		/// Represents Stacked Line Chart with data markers.
		/// </summary>
		[Description("Stacked line with markers")]
		StackedLineWithMarkers = 23,

		/// <summary>
		/// Represents 100% Stacked Line Chart with data markers.
		/// </summary>
		[Description("Percents stacked line with markers")]
		PercentsStackedLineWithMarkers = 24,

		/// <summary>
		/// Represents 3D Line Chart.
		/// </summary>
		[Description("Line 3D")]
		Line3D = 25,

		/// <summary>
		/// Represents Pie Chart.
		/// </summary>
		[Description("Pie")]
		Pie = 26,

		// TODO https://issue.lutsk.dynabic.com/issues/SLIDESAPP-561
		/// <summary>
		/// Represents 3D Pie Chart.
		/// </summary>
		//[Description("Pie 3D")] 
		//Pie3D = 27,

		/// <summary>
		/// Represents Pie of Pie Chart.
		/// </summary>
		[Description("Pie of pie")]
		PieOfPie = 28,

		/// <summary>
		/// Represents Exploded Pie Chart.
		/// </summary>
		[Description("Exploded pie")]
		ExplodedPie = 29,

		/// <summary>
		/// Represents 3D Exploded Pie Chart.
		/// </summary>
		[Description("Exploded pie 3D")]
		ExplodedPie3D = 30,

		/// <summary>
		/// Represents Bar of Pie Chart.
		/// </summary>
		[Description("Bar of pie")]
		BarOfPie = 31,

		/// <summary>
		/// Represents 100% Stacked Bar Chart.
		/// </summary>
		[Description("Percents stacked bar")]
		PercentsStackedBar = 32,

		/// <summary>
		/// Represents 3D Colustered Bar Chart.
		/// </summary>
		[Description("Clustered bar 3D")]
		ClusteredBar3D = 33,

		/// <summary>
		/// Represents Clustered Bar Chart.
		/// </summary>
		[Description("Clustered bar")]
		ClusteredBar = 34,

		/// <summary>
		/// Represents Stacked Bar Chart.
		/// </summary>
		[Description("Stacked bar")]
		StackedBar = 35,

		/// <summary>
		/// Represents 3D Stacked Bar Chart.
		/// </summary>
		[Description("Stacked bar 3D")]
		StackedBar3D = 36,

		/// <summary>
		/// Represents 3D 100% Stacked Bar Chart.
		/// </summary>
		[Description("Percents stacked bar 3D")]
		PercentsStackedBar3D = 37,

		/// <summary>
		/// Represents Cylindrical Bar Chart.
		/// </summary>
		[Description("Clustered horizontal cylinder")]
		ClusteredHorizontalCylinder = 38,

		/// <summary>
		/// Represents Stacked Cylindrical Bar Chart.
		/// </summary>
		[Description("Stacked horizontal cylinder")]
		StackedHorizontalCylinder = 39,

		/// <summary>
		/// Represents 100% Stacked Cylindrical Bar Chart.
		/// </summary>
		[Description("Percents stacked horizontal cylinder")]
		PercentsStackedHorizontalCylinder = 40,

		/// <summary>
		/// Represents Conical Bar Chart.
		/// </summary>
		[Description("Clustered horizontal cone")]
		ClusteredHorizontalCone = 41,

		/// <summary>
		/// Represents Stacked Conical Bar Chart.
		/// </summary>
		[Description("Stacked horizontal cone")]
		StackedHorizontalCone = 42,

		/// <summary>
		/// Represents 100% Stacked Conical Bar Chart.
		/// </summary>
		[Description("Percents stacked horizontal cone")]
		PercentsStackedHorizontalCone = 43,

		/// <summary>
		/// Represents Pyramid Bar Chart.
		/// </summary>
		[Description("Clustered horizontal pyramid")]
		ClusteredHorizontalPyramid = 44,

		/// <summary>
		/// Represents Stacked Pyramid Bar Chart.
		/// </summary>
		[Description("Stacked horizontal pyramid")]
		StackedHorizontalPyramid = 45,

		/// <summary>
		/// Represents 100% Stacked Pyramid Bar Chart.
		/// </summary>
		[Description("Percents stacked horizontal pyramid")]
		PercentsStackedHorizontalPyramid = 46,

		/// <summary>
		/// Represents Area Chart.
		/// </summary>
		[Description("Area")]
		Area = 47,

		/// <summary>
		/// Represents Stacked Area Chart.
		/// </summary>
		[Description("Stacked area")]
		StackedArea = 48,

		/// <summary>
		/// Represents 100% Stacked Area Chart.
		/// </summary>
		[Description("Percents stacked area")]
		PercentsStackedArea = 49,

		/// <summary>
		/// Represents 3D Area Chart.
		/// </summary>
		[Description("Area 3D")]
		Area3D = 50,

		/// <summary>
		/// Represents 3D Stacked Area Chart.
		/// </summary>
		[Description("Stacked area 3D")]
		StackedArea3D = 51,

		/// <summary>
		/// Represents 3D 100% Stacked Area Chart.
		/// </summary>
		[Description("Percents stacked area 3D")]
		PercentsStackedArea3D = 52,

		/// <summary>
		/// Represents Scatter Chart.
		/// </summary>
		[Description("Scatter with markers")]
		ScatterWithMarkers = 53,

		/// <summary>
		/// Represents Scatter Chart connected by curves, with data markers.
		/// </summary>
		[Description("Scatter with smooth lines and markers")]
		ScatterWithSmoothLinesAndMarkers = 54,

		/// <summary>
		/// Represents Scatter Chart connected by curves, without data markers.
		/// </summary>
		[Description("Scatter with smooth lines")]
		ScatterWithSmoothLines = 55,

		/// <summary>
		/// Represents Scatter Chart connected by lines, with data markers.
		/// </summary>
		[Description("Scatter with straight lines and markers")]
		ScatterWithStraightLinesAndMarkers = 56,

		/// <summary>
		/// Represents Scatter Chart connected by lines, without data markers.
		/// </summary>
		[Description("Scatter with straight lines")]
		ScatterWithStraightLines = 57,

		/// <summary>
		/// Represents High-Low-Close Stock Chart.
		/// </summary>
		[Description("High low close")]
		HighLowClose = 58,

		/// <summary>
		/// Represents Open-High-Low-Close Stock Chart.
		/// </summary>
		[Description("Open high low close")]
		OpenHighLowClose = 59,

		/// <summary>
		/// Represents Volume-High-Low-Close Stock Chart.
		/// </summary>
		[Description("Volume high low close")]
		VolumeHighLowClose = 60,

		/// <summary>
		/// Represents Volume-Open-High-Low-Close Stock Chart.
		/// </summary>
		[Description("Volume open high low close")]
		VolumeOpenHighLowClose = 61,

		/// <summary>
		/// Represents 3D Surface Chart.
		/// </summary>
		[Description("Surface 3D")]
		Surface3D = 62,

		/// <summary>
		/// Represents Wireframe 3D Surface Chart.
		/// </summary>
		[Description("Wireframe Surface 3D")]
		WireframeSurface3D = 63,

		/// <summary>
		/// Represents Contour Chart.
		/// </summary>
		[Description("Contour")]
		Contour = 64,

		/// <summary>
		/// Represents Wireframe Contour Chart.
		/// </summary>
		[Description("Wireframe Contour")]
		WireframeContour = 65,

		/// <summary>
		/// Represents Doughnut Chart.
		/// </summary>
		[Description("Doughnut")]
		Doughnut = 66,

		/// <summary>
		/// Represents Exploded Doughnut Chart.
		/// </summary>
		[Description("Exploded doughnut")]
		ExplodedDoughnut = 67,

		/// <summary>
		/// Represents Bubble Chart.
		/// </summary>
		[Description("Bubble")]
		Bubble = 68,

		/// <summary>
		/// Represents 3D Bubble Chart.
		/// </summary>
		[Description("Bubble with 3D")]
		BubbleWith3D = 69,

		/// <summary>
		/// Represents Radar Chart.
		/// </summary>
		[Description("Radar")]
		Radar = 70,

		/// <summary>
		/// Represents Radar Chart with data markers.
		/// </summary>
		[Description("Radar with markers")]
		RadarWithMarkers = 71,

		/// <summary>
		/// Represents Filled Radar Chart.
		/// </summary>
		[Description("Filled radar")]
		FilledRadar = 72,

		/// <summary>
		/// Represents Treemap chart.
		/// </summary>
		[Description("Treemap")]
		Treemap = 74,

		/// <summary>
		/// Represents Sunburst chart.
		/// </summary>
		[Description("Sunburst")]
		Sunburst = 75,

		/// <summary>
		/// Represents Histogram chart.
		/// </summary>
		[Description("Histogram")]
		Histogram = 76,

		/// <summary>
		/// Represents Pareto line series type (Histogram Pareto chart).
		/// </summary>
		[Description("Pareto line")]
		ParetoLine = 77,

		/// <summary>
		/// Represents BoxAndWhisker chart.
		/// </summary>
		[Description("Box and whisker")]
		BoxAndWhisker = 78,

		/// <summary>
		/// Represents Waterfall chart.
		/// </summary>
		[Description("Waterfall")]
		Waterfall = 79,

		/// <summary>
		/// Represents Funnel chart.
		/// </summary>
		[Description("Funnel")]
		Funnel = 80,

		// TODO: https://issue.lutsk.dynabic.com/issues/SLIDESAPP-555
		/// <summary>
		/// Represents Map chart.
		/// </summary>
		//[Description("Map")] 
		//Map = 81
	}
}
