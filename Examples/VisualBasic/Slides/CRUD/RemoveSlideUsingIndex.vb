Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.CRUD
    Public Class RemoveSlideUsingIndex
        Public Shared Sub Run()
            ' ExStart:RemoveSlideUsingIndex
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("RemoveSlideUsingIndex.pptx"))

                ' Removing a slide using its slide index
                pres.Slides.RemoveAt(0)

                ' Writing the presentation file
                pres.Save(dataDir & Convert.ToString("modified_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:RemoveSlideUsingIndex
        End Sub
    End Class
End Namespace
