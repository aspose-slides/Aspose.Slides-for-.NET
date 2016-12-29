Imports System
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class AddCustomError
        Public Shared Sub Run()
		    'ExStart:AddCustomError
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Creating empty presentation
            Using presentation As New Presentation()
                ' Creating a bubble chart
                Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.Bubble, 50, 50, 400, 300, True)

                ' Adding custom Error bars and setting its format
                Dim series As IChartSeries = chart.ChartData.Series(0)
                Dim errBarX As IErrorBarsFormat = series.ErrorBarsXFormat
                Dim errBarY As IErrorBarsFormat = series.ErrorBarsYFormat
                errBarX.IsVisible = True
                errBarY.IsVisible = True
                errBarX.ValueType = ErrorBarValueType.[Custom]
                errBarY.ValueType = ErrorBarValueType.[Custom]

                ' Accessing chart series data point and setting error bars values for individual point
                Dim points As IChartDataPointCollection = series.DataPoints
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForXPlusValues = DataSourceType.DoubleLiterals
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForXMinusValues = DataSourceType.DoubleLiterals
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForYPlusValues = DataSourceType.DoubleLiterals
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForYMinusValues = DataSourceType.DoubleLiterals

                ' Setting error bars for chart series points
                For i As Integer = 0 To points.Count - 1
                    points(i).ErrorBarsCustomValues.XMinus.AsLiteralDouble = i + 1
                    points(i).ErrorBarsCustomValues.XPlus.AsLiteralDouble = i + 1
                    points(i).ErrorBarsCustomValues.YMinus.AsLiteralDouble = i + 1
                    points(i).ErrorBarsCustomValues.YPlus.AsLiteralDouble = i + 1
                Next

                ' Saving presentation
                presentation.Save(dataDir & Convert.ToString("ErrorBarsCustomValues_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:AddCustomError
        End Sub
    End Class
End Namespace