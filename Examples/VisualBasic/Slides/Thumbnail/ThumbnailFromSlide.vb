Imports System
Imports System.Drawing
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Thumbnail
    Public Class ThumbnailFromSlide
        Public Shared Sub Run()
            ' ExStart:ThumbnailFromSlide
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Thumbnail()

            ' Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("ThumbnailFromSlide.pptx"))

                ' Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Create a full scale image
                Dim bmp As Bitmap = sld.GetThumbnail(1.0F, 1.0F)

                ' Save the image to disk in JPEG format
                bmp.Save(dataDir & Convert.ToString("Thumbnail_out.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg)
            End Using
            ' ExEnd:ThumbnailFromSlide
        End Sub
    End Class
End Namespace