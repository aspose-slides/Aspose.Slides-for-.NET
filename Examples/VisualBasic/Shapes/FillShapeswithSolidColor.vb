Imports System
Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FillShapeswithSolidColor
        Public Shared Sub Run()
			'ExStart:FillShapeswithSolidColor		
            ' The path to the documents directory.                    
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get the first slide
            Dim slide As ISlide = presentation.Slides(0)

            ' Add autoshape of rectangle type
            Dim shape As IShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)

            ' Set the fill type to Solid
            shape.FillFormat.FillType = FillType.Solid

            ' Set the color of the rectangle
            shape.FillFormat.SolidFillColor.Color = Color.Yellow

            ' Write the PPTX file to disk
            presentation.Save(dataDir & Convert.ToString("RectShpSolid_out.pptx"), SaveFormat.Pptx)
			'ExEnd:FillShapeswithSolidColor		
        End Sub
    End Class
End Namespace
