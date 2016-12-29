Imports System
Imports System.Drawing

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Thumbnail
    Public Class ThumbnailWithUserDefinedDimensions
        Public Shared Sub Run()
            ' ExStart:ThumbnailWithUserDefinedDimensions

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Thumbnail()

            ' Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("ThumbnailWithUserDefinedDimensions.pptx"))

                ' Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' User defined dimension
                Dim desiredX As Integer = 1200
                Dim desiredY As Integer = 800

                ' Getting scaled value  of X and Y
                Dim ScaleX As Single = CSng(1.0 / pres.SlideSize.Size.Width) * desiredX
                Dim ScaleY As Single = CSng(1.0 / pres.SlideSize.Size.Height) * desiredY

                ' Create a full scale image
                Dim bmp As Bitmap = sld.GetThumbnail(ScaleX, ScaleY)

                ' Save the image to disk in JPEG format
                bmp.Save(dataDir & Convert.ToString("Thumbnail2_out.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg)
            End Using
            ' ExEnd:ThumbnailWithUserDefinedDimensions
        End Sub
    End Class
End Namespace