Imports System
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class SetDataRange
        Public Shared Sub Run()
			'ExStart:SetDataRange
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Instantiate Presentation class that represents PPTX file
            Dim presentation As New Presentation(dataDir & Convert.ToString("ExistingChart.pptx"))

            ' Access first slideMarker and add chart with default data
            Dim slide As ISlide = presentation.Slides(0)
            Dim chart As IChart = DirectCast(slide.Shapes(0), IChart)
            chart.ChartData.SetRange("Sheet1!A1:B4")
            presentation.Save(dataDir & Convert.ToString("SetDataRange_out.pptx"), SaveFormat.Pptx)
        End Sub
			'ExEnd:SetDataRange
    End Class
End Namespace