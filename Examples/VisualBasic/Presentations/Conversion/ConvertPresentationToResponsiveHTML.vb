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
    Public Class ConvertPresentationToResponsiveHTML
        Public Shared Sub Run()
			'ExStart:ConvertPresentationToResponsiveHTML
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("Convert_HTML.pptx"))
                Dim controller As New ResponsiveHtmlController()
                Dim htmlOptions As New HtmlOptions() With { _
                     .HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller) _
                }

                ' Saving the presentation to HTML
                presentation.Save(dataDir & Convert.ToString("ConvertPresentationToResponsiveHTML_out.html"), SaveFormat.Html, HtmlOptions)
            End Using
			'ExEnd:ConvertPresentationToResponsiveHTML
        End Sub
    End Class
End Namespace
