Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Saving
    Public Class SaveToFile
        Public Shared Sub Run()
			'ExStart:SaveToFile
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationSaving()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If Not IsExists Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a Presentation object that represents a PPT file
            Dim presentation As New Presentation()

            '...do some work here...

            ' Save your presentation to a file
            presentation.Save(dataDir & Convert.ToString("Saved_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
			'ExEnd:SaveToFile
        End Sub
    End Class
End Namespace
