Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Conversion
    Public Class ConvertWholePresentationToHTML
        Public Shared Sub Run()
			'ExStart:ConvertToPDFWithHiddenSlides
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("Convert_HTML.pptx"))
                Dim htmlOpt As New HtmlOptions()
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", False)

                ' Saving the presentation to HTML
                presentation.Save(dataDir & Convert.ToString("ConvertWholePresentationToHTML_out.html"), SaveFormat.Html, htmlOpt)
            End Using
			'ExEnd:ConvertToPDFWithHiddenSlides
        End Sub
    End Class
End Namespace
