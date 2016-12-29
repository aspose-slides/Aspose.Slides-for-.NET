Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class SetAlternativeText
        Public Shared Sub Run()
			'ExStart:SetAlternativeText	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX
            Dim pres As New Presentation()

            ' Get the first slide
            Dim sld As ISlide = pres.Slides(0)

            ' Add autoshape of rectangle type
            Dim shp1 As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 40, 150, 50)
            Dim shp2 As IShape = sld.Shapes.AddAutoShape(ShapeType.Moon, 160, 40, 150, 50)
            shp2.FillFormat.FillType = FillType.Solid
            shp2.FillFormat.SolidFillColor.Color = Color.Gray

            For i As Integer = 0 To sld.Shapes.Count - 1
                Dim autoShape As AutoShape = TryCast(sld.Shapes(i), AutoShape)
                If (autoShape IsNot Nothing) Then
                    autoShape.AlternativeText = "User Defined"
                End If
            Next

            ' Save presentation to disk
            pres.Save(dataDir + "Set_AlternativeText_out.pptx", SaveFormat.Pptx)
			'ExEnd:SetAlternativeText	
        End Sub
    End Class
End Namespace