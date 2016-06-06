'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Charts
    Public Class NumberFormat
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate the presentation//Instantiate the presentation
            Dim pres As New Presentation()

            'Access the first presentation slide
            Dim slide As ISlide = pres.Slides(0)

            'Adding a defautlt clustered column chart
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400)

            'Accessing the chart series collection
            Dim series As IChartSeriesCollection = chart.ChartData.Series

            'Setting the preset number format
            'Traverse through every chart series
            For Each ser As ChartSeries In series
                'Traverse through every data cell in series
                For Each cell As IChartDataPoint In ser.DataPoints
                    'Setting the number format
                    cell.Value.AsCell.PresetNumberFormat = 10 '0.00%
                Next cell
            Next ser

            'Saving presentation
            pres.Save(dataDir & "PresetNumberFormat.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace