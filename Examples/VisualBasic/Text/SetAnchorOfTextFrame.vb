Imports System
Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class SetAnchorOfTextFrame
        Public Shared Sub Run()
            ' ExStart:SetAnchorOfTextFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get the first slide 
            Dim slide As ISlide = presentation.Slides(0)

            ' Add an AutoShape of Rectangle type
            Dim ashp As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 350, 350)

            ' Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ")
            ashp.FillFormat.FillType = FillType.NoFill

            ' Accessing the text frame
            Dim txtFrame As ITextFrame = ashp.TextFrame
            txtFrame.TextFrameFormat.AnchoringType = TextAnchorType.Bottom

            ' Create the Paragraph object for text frame
            Dim para As IParagraph = txtFrame.Paragraphs(0)

            ' Create Portion object for paragraph
            Dim portion As IPortion = para.Portions(0)
            portion.Text = "A quick brown fox jumps over the lazy dog. A quick brown fox jumps over the lazy dog."
            portion.PortionFormat.FillFormat.FillType = FillType.Solid
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black

            ' Save Presentation
            presentation.Save(dataDir & Convert.ToString("AnchorText_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:SetAnchorOfTextFrame

        End Sub
    End Class
End Namespace