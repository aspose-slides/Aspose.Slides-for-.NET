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
    Public Class AddErrorBars
        Public Shared Sub Run()
			'ExStart:AddErrorBars
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Creating empty presentation
            Using presentation As New Presentation()
                ' Creating a bubble chart
                Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.Bubble, 50, 50, 400, 300, True)

                ' Adding Error bars and setting its format
                Dim errBarX As IErrorBarsFormat = chart.ChartData.Series(0).ErrorBarsXFormat
                Dim errBarY As IErrorBarsFormat = chart.ChartData.Series(0).ErrorBarsYFormat
                errBarX.IsVisible = True
                errBarY.IsVisible = True
                errBarX.ValueType = ErrorBarValueType.Fixed
                errBarX.Value = 0.1F
                errBarY.ValueType = ErrorBarValueType.Percentage
                errBarY.Value = 5
                errBarX.Type = ErrorBarType.Plus
                errBarY.Format.Line.Width = 2
                errBarX.HasEndCap = True

                ' Saving presentation
                presentation.Save(dataDir & Convert.ToString("ErrorBars_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:AddErrorBars
        End Sub
    End Class
End Namespace