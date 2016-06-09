Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace VisualBasic.Slides
    Class SetPDFPageSize
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate a Presentation object that represents a presentation file 
            Dim presentation As New Presentation()

            ' Set SlideSize.Type Property 
            presentation.SlideSize.Type = SlideSizeType.A4Paper

            ' Set different properties of PDF Options
            Dim opts As New PdfOptions()
            opts.SufficientResolution = 600

            ' Save presentation to disk
            presentation.Save(dataDir & Convert.ToString("SetPDFPageSize.pdf"), SaveFormat.Pdf, opts)
        End Sub
    End Class
End Namespace