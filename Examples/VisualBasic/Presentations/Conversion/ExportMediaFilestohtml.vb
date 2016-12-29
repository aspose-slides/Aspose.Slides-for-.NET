Imports System
Imports System.IO
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Conversion
    Class ExportMediaFilestohtml
        Public Shared Sub Run()
			'ExStart:ExportMediaFilestohtml
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Loading a presentation
            Using pres As New Presentation(dataDir & Convert.ToString("Media File.pptx"))
                Dim path__1 As String = dataDir
                Const fileName As String = "ExportMediaFiles_out.html"
                Const baseUri As String = "http://www.example.com/"

                Dim controller As New VideoPlayerHtmlController(path__1, fileName, baseUri)

                ' Setting HTML options
                Dim htmlOptions As New HtmlOptions(controller)
                Dim svgOptions As New SVGOptions(controller)

                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller)
                htmlOptions.SlideImageFormat = SlideImageFormat.Svg(svgOptions)

                ' Saving the file
                pres.Save(Path.Combine(path__1, fileName), SaveFormat.Html, htmlOptions)
            End Using
			'ExEnd:ExportMediaFilestohtml
        End Sub
    End Class
End Namespace
