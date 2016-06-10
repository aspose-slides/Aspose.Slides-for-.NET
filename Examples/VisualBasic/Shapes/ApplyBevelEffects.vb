'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System.Drawing
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class ApplyBevelEffects
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create an instance of Presentation class
            Dim pres As New Presentation()
            Dim slide As ISlide = pres.Slides(0)

            ' Add a shape on slide
            Dim shape As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Ellipse, 30, 30, 100, 100)
            shape.FillFormat.FillType = FillType.Solid
            shape.FillFormat.SolidFillColor.Color = Color.Green
            Dim format As ILineFillFormat = shape.LineFormat.FillFormat
            format.FillType = FillType.Solid
            format.SolidFillColor.Color = Color.Orange
            shape.LineFormat.Width = 2.0

            ' Set ThreeDFormat properties of shape
            shape.ThreeDFormat.Depth = 4
            shape.ThreeDFormat.BevelTop.BevelType = BevelPresetType.Circle
            shape.ThreeDFormat.BevelTop.Height = 6
            shape.ThreeDFormat.BevelTop.Width = 6
            shape.ThreeDFormat.Camera.CameraType = CameraPresetType.OrthographicFront
            shape.ThreeDFormat.LightRig.LightType = LightRigPresetType.ThreePt
            shape.ThreeDFormat.LightRig.Direction = LightingDirection.Top

            ' Write the presentation as a PPTX file
            pres.Save(dataDir + "Bavel.pptx", SaveFormat.Pptx)

        End Sub
    End Class
End Namespace










