Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Class ChangeSmartArtState
        Public Shared Sub Run()
            ' ExStart:ChangeSmartArtState
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            Using presentation As New Presentation()
                ' Add SmartArt BasicProcess 
                Dim smart As ISmartArt = presentation.Slides(0).Shapes.AddSmartArt(10, 10, 400, 300, SmartArtLayoutType.BasicProcess)

                ' Get or Set the state of SmartArt Diagram
                smart.IsReversed = True
                Dim flag As Boolean = smart.IsReversed
                ' Saving Presentation
                presentation.Save(dataDir & Convert.ToString("ChangeSmartArtState_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:ChangeSmartArtState
        End Sub
    End Class
End Namespace