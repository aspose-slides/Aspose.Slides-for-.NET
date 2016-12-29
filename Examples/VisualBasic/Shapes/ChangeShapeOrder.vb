Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Class ChangeShapeOrder
        Public Shared Sub Run()

			'ExStart:ChangeShapeOrder	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents PPTX file
            Dim presentation1 As New Presentation(dataDir & Convert.ToString("HelloWorld.pptx"))

            ' Get the first slide
            Dim slide As ISlide = presentation1.Slides(0)

            ' Adding rectangle shape in slide
            Dim shp3 As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 365, 400, 150)
            shp3.FillFormat.FillType = FillType.NoFill
            shp3.AddTextFrame(" ")

            ' Adding text in slide
            Dim txtFrame As ITextFrame = shp3.TextFrame
            Dim para As IParagraph = txtFrame.Paragraphs(0)
            Dim portion As IPortion = para.Portions(0)
            portion.Text = "Watermark Text Watermark Text Watermark Text"
            shp3 = slide.Shapes.AddAutoShape(ShapeType.Triangle, 200, 365, 400, 150)
            slide.Shapes.Reorder(2, shp3)

            ' Save presentation
            presentation1.Save(dataDir & Convert.ToString("Reshape_out.pptx"), SaveFormat.Pptx)
			'ExEnd:ChangeShapeOrder
        End Sub
    End Class
End Namespace
