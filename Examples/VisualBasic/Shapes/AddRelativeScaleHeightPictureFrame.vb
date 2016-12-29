Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddRelativeScaleHeightPictureFrame
        Public Shared Sub Run()
			'ExStart:AddRelativeScaleHeightPictureFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class object
            Using prsentation As New Presentation()

                ' Get the first slide
                Dim slide As ISlide = prsentation.Slides(0)

                ' Instantiate the ImageEx class
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir + "aspose-logo.jpg"), System.Drawing.Image)
                Dim image As IPPImage = prsentation.Images.AddImage(img)

                ' Add Picture Frame with height and width equivalent of Picture
                Dim pf As IPictureFrame = slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image)

                ' Setting relative scale width and height
                pf.RelativeScaleHeight = 0.8F
                pf.RelativeScaleWidth = 1.35F

                ' Write the PPTX file to disk
                prsentation.Save(dataDir + "Adding Picture Frame with Relative Scale_out.pptx", SaveFormat.Pptx)

            End Using
			'ExEnd:AddRelativeScaleHeightPictureFrame
        End Sub
    End Class
End Namespace