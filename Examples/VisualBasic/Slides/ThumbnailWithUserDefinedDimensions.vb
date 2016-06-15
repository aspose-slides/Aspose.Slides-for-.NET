Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class ThumbnailWithUserDefinedDimensions
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & "ThumbnailWithUserDefinedDimensions.pptx")

                'Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                'User defined dimension
                Dim desiredX As Integer = 1200
                Dim desiredY As Integer = 800

                'Getting scaled value  of X and Y
                Dim ScaleX As Single = CSng(1.0 / pres.SlideSize.Size.Width) * desiredX
                Dim ScaleY As Single = CSng(1.0 / pres.SlideSize.Size.Height) * desiredY


                'Create a full scale image
                Dim bmp As Bitmap = sld.GetThumbnail(ScaleX, ScaleY)

                'Save the image to disk in JPEG format
                bmp.Save(dataDir & "Thumbnail2.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

            End Using
        End Sub


    End Class
End Namespace

