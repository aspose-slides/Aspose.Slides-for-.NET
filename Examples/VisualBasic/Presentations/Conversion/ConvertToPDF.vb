Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Conversion
    Public Class ConvertToPDF
        Public Shared Sub Run()
			'ExStart:ConvertToPDF
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Dim presentation As New Presentation(dataDir & Convert.ToString("ConvertToPDF.pptx"))

            ' Save the presentation to PDF with default options
            presentation.Save(dataDir & Convert.ToString("output_out.pdf"), SaveFormat.Pdf)
			'ExStart:ConvertToPDF
        End Sub
    End Class
End Namespace
