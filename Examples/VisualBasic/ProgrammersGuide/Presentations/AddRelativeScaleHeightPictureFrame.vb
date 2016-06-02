'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.Drawing
Imports System.IO
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class AddRelativeScaleHeightPictureFrame
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            'Instantiate Prseetation class object
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Instantiate the ImageEx class
                Dim img As System.Drawing.Image = CType(New Bitmap(dataDir + "aspose-logo.jpg"), System.Drawing.Image)
                Dim image As IPPImage = pres.Images.AddImage(img)

                'Add Picture Frame with height and width equivalent of Picture
                Dim pf As IPictureFrame = sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image)

                'etting relative scale width and height
                pf.RelativeScaleHeight = 0.8F
                pf.RelativeScaleWidth = 1.35F

                'Write the PPTX file to disk
                pres.Save(dataDir + "Adding Picture Frame with Relative Scale.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace