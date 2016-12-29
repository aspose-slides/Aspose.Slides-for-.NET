Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Background
    Public Class SetBackgroundToGradient
        Public Shared Sub Run()
            'ExStart:SetBackgroundToGradient

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Background()

            ' Instantiate the Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("SetBackgroundToGradient.pptx"))

                ' Apply Gradiant effect to the Background
                pres.Slides(0).Background.Type = BackgroundType.OwnBackground
                pres.Slides(0).Background.FillFormat.FillType = FillType.Gradient
                pres.Slides(0).Background.FillFormat.GradientFormat.TileFlip = TileFlip.FlipBoth

                'Write the presentation to disk
                pres.Save(dataDir & Convert.ToString("ContentBG_Grad_out.pptx"), SaveFormat.Pptx)
            End Using
            'ExEnd:SetBackgroundToGradient
        End Sub
    End Class
End Namespace
