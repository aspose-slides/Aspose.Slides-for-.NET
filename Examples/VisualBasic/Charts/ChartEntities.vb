Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides.Charts

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class ChartEntities
        Public Shared Sub Run()
			'ExStart:ChartEntities
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiating presentation//Instantiating presentation
            Dim pres As New Presentation()

            ' Accessing the first slide
            Dim slide As ISlide = pres.Slides(0)

            ' Adding the sample chart
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 50, 50, 500, 400)

            ' Setting Chart Titile
            chart.HasTitle = True
            chart.ChartTitle.AddTextFrameForOverriding("")
            Dim chartTitle As IPortion = chart.ChartTitle.TextFrameForOverriding.Paragraphs(0).Portions(0)
            chartTitle.Text = "Sample Chart"
            chartTitle.PortionFormat.FillFormat.FillType = FillType.Solid
            chartTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
            chartTitle.PortionFormat.FontHeight = 20
            chartTitle.PortionFormat.FontBold = NullableBool.True
            chartTitle.PortionFormat.FontItalic = NullableBool.True

            ' Setting Major grid lines format for value axis
            chart.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid
            chart.Axes.VerticalAxis.MajorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Blue
            chart.Axes.VerticalAxis.MajorGridLinesFormat.Line.Width = 5
            chart.Axes.VerticalAxis.MajorGridLinesFormat.Line.DashStyle = LineDashStyle.DashDot

            ' Setting Minor grid lines format for value axis
            chart.Axes.VerticalAxis.MinorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid
            chart.Axes.VerticalAxis.MinorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Red
            chart.Axes.VerticalAxis.MinorGridLinesFormat.Line.Width = 3

            ' Setting value axis number format
            chart.Axes.VerticalAxis.IsNumberFormatLinkedToSource = False
            chart.Axes.VerticalAxis.DisplayUnit = DisplayUnitType.Thousands
            chart.Axes.VerticalAxis.NumberFormat = "0.0%"

            ' Setting chart maximum, minimum values
            chart.Axes.VerticalAxis.IsAutomaticMajorUnit = False
            chart.Axes.VerticalAxis.IsAutomaticMaxValue = False
            chart.Axes.VerticalAxis.IsAutomaticMinorUnit = False
            chart.Axes.VerticalAxis.IsAutomaticMinValue = False

            chart.Axes.VerticalAxis.MaxValue = 15.0F
            chart.Axes.VerticalAxis.MinValue = -2.0F
            chart.Axes.VerticalAxis.MinorUnit = 0.5F
            chart.Axes.VerticalAxis.MajorUnit = 2.0F

            ' Setting Value Axis Text Properties
            Dim txtVal As IChartPortionFormat = chart.Axes.VerticalAxis.TextFormat.PortionFormat
            txtVal.FontBold = NullableBool.True
            txtVal.FontHeight = 16
            txtVal.FontItalic = NullableBool.True
            txtVal.FillFormat.FillType = FillType.Solid

            txtVal.FillFormat.SolidFillColor.Color = Color.DarkGreen
            txtVal.LatinFont = New FontData("Times New Roman")

            ' Setting value axis title
            chart.Axes.VerticalAxis.HasTitle = True
            chart.Axes.VerticalAxis.Title.AddTextFrameForOverriding("")
            Dim valtitle As IPortion = chart.Axes.VerticalAxis.Title.TextFrameForOverriding.Paragraphs(0).Portions(0)
            valtitle.Text = "Primary Axis"
            valtitle.PortionFormat.FillFormat.FillType = FillType.Solid
            valtitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
            valtitle.PortionFormat.FontHeight = 20
            valtitle.PortionFormat.FontBold = NullableBool.True
            valtitle.PortionFormat.FontItalic = NullableBool.True

            ' Setting value axis line format : Now Obselete
            ' chart.Axes.VerticalAxis.aVerticalAxis.l.AxisLine.Width = 10
            ' chart.Axes.VerticalAxis.AxisLine.FillFormat.FillType = FillType.Solid
            ' Chart.Axes.VerticalAxis.AxisLine.FillFormat.SolidFillColor.Color = Color.Red

            ' Setting Major grid lines format for Category axis
            chart.Axes.HorizontalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid
            chart.Axes.HorizontalAxis.MajorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Green
            chart.Axes.HorizontalAxis.MajorGridLinesFormat.Line.Width = 5

            ' Setting Minor grid lines format for Category axis
            chart.Axes.HorizontalAxis.MinorGridLinesFormat.Line.FillFormat.FillType = FillType.Solid
            chart.Axes.HorizontalAxis.MinorGridLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.Yellow
            chart.Axes.HorizontalAxis.MinorGridLinesFormat.Line.Width = 3

            ' Setting Category Axis Text Properties
            Dim txtCat As IChartPortionFormat = chart.Axes.HorizontalAxis.TextFormat.PortionFormat
            txtCat.FontBold = NullableBool.True
            txtCat.FontHeight = 16
            txtCat.FontItalic = NullableBool.True
            txtCat.FillFormat.FillType = FillType.Solid

            txtCat.FillFormat.SolidFillColor.Color = Color.Blue
            txtCat.LatinFont = New FontData("Arial")

            ' Setting Category Titile
            chart.Axes.HorizontalAxis.HasTitle = True
            chart.Axes.HorizontalAxis.Title.AddTextFrameForOverriding("")

            Dim catTitle As IPortion = chart.Axes.HorizontalAxis.Title.TextFrameForOverriding.Paragraphs(0).Portions(0)
            catTitle.Text = "Sample Category"
            catTitle.PortionFormat.FillFormat.FillType = FillType.Solid
            catTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
            catTitle.PortionFormat.FontHeight = 20
            catTitle.PortionFormat.FontBold = NullableBool.True
            catTitle.PortionFormat.FontItalic = NullableBool.True

            ' Setting category axis lable position
            chart.Axes.HorizontalAxis.TickLabelPosition = TickLabelPositionType.Low

            ' Setting category axis lable rotation angle
            chart.Axes.HorizontalAxis.TickLabelRotationAngle = 45

            ' Setting Legends Text Properties
            Dim txtleg As IChartPortionFormat = chart.Legend.TextFormat.PortionFormat
            txtleg.FontBold = NullableBool.True
            txtleg.FontHeight = 16
            txtleg.FontItalic = NullableBool.True
            txtleg.FillFormat.FillType = FillType.Solid

            txtleg.FillFormat.SolidFillColor.Color = Color.DarkRed

            ' Set show chart legends without overlapping chart

            chart.Legend.Overlay = True

            ' Ploting first series on secondary value axis
            ' Chart.ChartData.Series[0].PlotOnSecondAxis = true

            ' Setting chart back wall color
            chart.BackWall.Thickness = 1
            chart.BackWall.Format.Fill.FillType = FillType.Solid
            chart.BackWall.Format.Fill.SolidFillColor.Color = Color.Orange

            chart.Floor.Format.Fill.FillType = FillType.Solid
            chart.Floor.Format.Fill.SolidFillColor.Color = Color.Red
            ' Setting Plot area color
            chart.PlotArea.Format.Fill.FillType = FillType.Solid
            chart.PlotArea.Format.Fill.SolidFillColor.Color = Color.LightCyan

            ' Save Presentation
            pres.Save(dataDir & "FormattedChart_out.pptx", SaveFormat.Pptx)

			'ExEnd:ChartEntities
        End Sub
    End Class
End Namespace