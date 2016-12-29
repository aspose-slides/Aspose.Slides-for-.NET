Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.CRUD
    Public Class ChangePosition
        Public Shared Sub Run()
            'ExStart:ChangePosition
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Instantiate Presentation class to load the source presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("ChangePosition.pptx"))
                ' Get the slide whose position is to be changed
                Dim sld As ISlide = pres.Slides(0)

                ' Set the new position for the slide
                sld.SlideNumber = 2

                ' Write the presentation to disk
                pres.Save(dataDir & Convert.ToString("ChangePosition_out.pptx"), SaveFormat.Pptx)
            End Using
            'ExEnd:ChangePosition
        End Sub
    End Class
End Namespace
