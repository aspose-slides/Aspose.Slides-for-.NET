Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddSimplePictureFrames
        Public Shared Sub Run()
			'ExStart:AddSimplePictureFrames
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Instantiate the ImageEx class
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "aspose-logo.jpg"), System.Drawing.Image)
                Dim imgx As IPPImage = pres.Images.AddImage(img)

                ' Add Picture Frame with height and width equivalent of Picture
                sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx)

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectPicFrame_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:AddSimplePictureFrames
        End Sub
    End Class
End Namespace