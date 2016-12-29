Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.CRUD
    Public Class CloneWithinSamePresentationToEnd
        Public Shared Sub Run()
            ' ExStart:CloneWithinSamePresentationToEnd
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Instantiate Presentation class that represents a presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("CloneWithinSamePresentationToEnd.pptx"))

                ' Clone the desired slide to the end of the collection of slides in the same presentation
                Dim slds As ISlideCollection = pres.Slides

                slds.AddClone(pres.Slides(0))

                ' Write the modified presentation to disk
                pres.Save(dataDir & Convert.ToString("Aspose_CloneWithinSamePresentationToEnd_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:CloneWithinSamePresentationToEnd
        End Sub
    End Class
End Namespace
