Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Conversion
    Public Class ConvertWithoutXpsOptions
        Public Shared Sub Run()
			'ExStart:ConvertWithoutXpsOptions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Conversion()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("Convert_XPS.pptx"))
                ' Saving the presentation to XPS document
                pres.Save(dataDir & Convert.ToString("XPS_Output_Without_XPSOption_out.xps"), SaveFormat.Xps)
            End Using
			'ExEnd:ConvertWithoutXpsOptions
        End Sub
    End Class
End Namespace
