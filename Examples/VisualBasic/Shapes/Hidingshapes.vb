Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class Hidingshapes
        Public Shared Sub Run()
			'ExStart:Hidingshapes	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX
            Dim presentation As New Presentation()

            ' Get the first slide
            Dim sld As ISlide = presentation.Slides(0)

            ' Add autoshape of rectangle type
            Dim shp1 As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 40, 150, 50)
            Dim shp2 As IShape = sld.Shapes.AddAutoShape(ShapeType.Moon, 160, 40, 150, 50)
            Dim alttext As [String] = "User Defined"
            Dim iCount As Integer = sld.Shapes.Count
            For i As Integer = 0 To iCount - 1
                Dim ashp As AutoShape = DirectCast(sld.Shapes(i), AutoShape)
                If String.Compare(ashp.AlternativeText, alttext, StringComparison.Ordinal) = 0 Then
                    ashp.Hidden = True
                End If
            Next

            ' Save presentation to disk
            presentation.Save(dataDir + "Hidding_shapes_out.pptx", SaveFormat.Pptx)
			'ExEnd:Hidingshapes	
        End Sub
    End Class
End Namespace