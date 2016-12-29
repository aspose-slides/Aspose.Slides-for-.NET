Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx 
' 

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Conversion
    Public Class ConvertWithXpsOptions
        Public Shared Sub Run()
			'ExStart:ConvertWithXpsOptions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("Convert_XPS_Options.pptx"))
                ' Instantiate the TiffOptions class
                Dim opts As New XpsOptions()

                ' Save MetaFiles as PNG
                opts.SaveMetafilesAsPng = True

                ' Save the presentation to XPS document
                pres.Save(dataDir & Convert.ToString("XPS_With_Options_out.xps"), SaveFormat.Xps, opts)
            End Using
			'ExEnd:ConvertWithXpsOptions
        End Sub
    End Class
End Namespace
