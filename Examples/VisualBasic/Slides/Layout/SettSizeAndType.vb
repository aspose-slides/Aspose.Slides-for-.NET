Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Layout
    Class SetSizeAndType
        Public Shared Sub Run()
            ' ExStart:SetSizeAndType
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Layout()

            ' Instantiate a Presentation object that represents a presentation file 
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))
            Dim auxPresentation As New Presentation()

            Dim slide As ISlide = presentation.Slides(0)

            ' Set the slide size of generated presentations to that of source
            auxPresentation.SlideSize.Type = presentation.SlideSize.Type
            auxPresentation.SlideSize.Size = presentation.SlideSize.Size

            auxPresentation.Slides.InsertClone(0, slide)
            auxPresentation.Slides.RemoveAt(0)
            ' Save Presentation to disk
            auxPresentation.Save(dataDir & Convert.ToString("Set_Size&Type_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:SettSizeAndType
        End Sub
    End Class
End Namespace