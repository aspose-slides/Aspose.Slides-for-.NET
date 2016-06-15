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

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Class CloneAnotherPresentationAtSpecifiedPosition
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' ExStart:CloneAnotherPresentationAtSpecifiedPosition
            ' Instantiate Presentation class to load the source presentation file
            Using sourcePresentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))
                ' Instantiate Presentation class for destination presentation (where slide is to be cloned)
                Using destPres As New Presentation()
                    ' Clone the desired slide from the source presentation to the end of the collection of slides in destination presentation
                    Dim slideCollection As ISlideCollection = destPres.Slides

                    ' Clone the desired slide from the source presentation to the specified position in destination presentation
                    slideCollection.InsertClone(1, sourcePresentation.Slides(1))

                    ' ExEnd:CloneAnotherPresentationAtSpecifiedPosition
                    ' Write the destination presentation to disk
                    destPres.Save(dataDir & Convert.ToString("CloneAnotherPresentationAtSpecifiedPosition.pptx"), SaveFormat.Pptx)
                End Using
            End Using
        End Sub
    End Class
End Namespace