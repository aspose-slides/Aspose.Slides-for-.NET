'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System.Drawing
Imports Aspose.Slides
Imports VisualBasic

Namespace ProgrammersGuide.Presentations
    Public Class ConvertSlidesToPdfNotes
        Public Shared Sub Run()

            Dim dataDir As String = RunExamples.GetDataDir_Presentations()
           
            'Instantiate a Presentation object that represents a presentation file 
            Dim presentation As Presentation = New Presentation(dataDir & "SelectedSlides.pptx")

            Dim auxPresentation As Presentation = New Presentation()
            Dim slide As ISlide = presentation.Slides(0)

            auxPresentation.Slides.InsertClone(0, slide)
            auxPresentation.SlideSize.Type = SlideSizeType.[Custom]
            auxPresentation.SlideSize.Size = New SizeF(612.0F, 792.0F)

            'Save Presentation to disk
            auxPresentation.Save(dataDir & "Converted-PDFNotes.pdf", Export.SaveFormat.Pdf)

        End Sub
    End Class
End Namespace