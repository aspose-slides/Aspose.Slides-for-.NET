Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class CreateGroupShape
        Public Shared Sub Run()
			'ExStart:CreateGroupShape
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class 
            Using pres As New Presentation()

                ' Get the first slide 
                Dim sld As ISlide = pres.Slides(0)

                ' Accessing the shape collection of slides 
                Dim slideShapes As IShapeCollection = sld.Shapes

                ' Adding a group shape to the slide 
                Dim groupShape As IGroupShape = slideShapes.AddGroupShape()

                ' Adding shapes inside added group shape 
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 100, 100, 100)
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 500, 100, 100, 100)
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 300, 100, 100)
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 500, 300, 100, 100)

                ' Adding group shape frame 
                groupShape.Frame = New ShapeFrame(100, 300, 500, 40, NullableBool.False, NullableBool.False, 0)

                ' Write the PPTX file to disk 
                pres.Save(dataDir + "GroupShape_out.pptx", SaveFormat.Pptx)

            End Using
			'ExStart:CreateGroupShape
        End Sub
    End Class
End Namespace