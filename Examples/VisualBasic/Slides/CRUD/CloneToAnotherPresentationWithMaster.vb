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
    Public Class CloneToAnotherPresentationWithMaster
        Public Shared Sub Run()
            ' ExStart:CloneToAnotherPresentationWithMaster
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Instantiate Presentation class to load the source presentation file

            Using srcPres As New Presentation(dataDir & Convert.ToString("CloneToAnotherPresentationWithMaster.pptx"))
                ' Instantiate Presentation class for destination presentation (where slide is to be cloned)
                Using destPres As New Presentation()

                    ' Instantiate ISlide from the collection of slides in source presentation along with
                    ' Master slide
                    Dim SourceSlide As ISlide = srcPres.Slides(0)
                    Dim SourceMaster As IMasterSlide = SourceSlide.LayoutSlide.MasterSlide

                    ' Clone the desired master slide from the source presentation to the collection of masters in the
                    ' Destination presentation
                    Dim masters As IMasterSlideCollection = destPres.Masters
                    Dim DestMaster As IMasterSlide = SourceSlide.LayoutSlide.MasterSlide

                    ' Clone the desired master slide from the source presentation to the collection of masters in the
                    ' Destination presentation
                    Dim iSlide As IMasterSlide = masters.AddClone(SourceMaster)

                    ' Clone the desired slide from the source presentation with the desired master to the end of the
                    ' Collection of slides in the destination presentation
                    Dim slds As ISlideCollection = destPres.Slides
                    slds.AddClone(SourceSlide, iSlide, True)

                    ' Clone the desired master slide from the source presentation to the collection of masters in the // Destination presentation
                    ' Save the destination presentation to disk

                    destPres.Save(dataDir & Convert.ToString("CloneToAnotherPresentationWithMaster_out.pptx"), SaveFormat.Pptx)
                End Using
            End Using
            ' ExEnd:CloneToAnotherPresentationWithMaster
        End Sub
    End Class
End Namespace
