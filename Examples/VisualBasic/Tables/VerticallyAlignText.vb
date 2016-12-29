Imports System.Drawing
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Tables
    Public Class VerticallyAlignText
        Public Shared Sub Run()
            ' ExStart:VerticallyAlignText
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Tables()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get the first slide 
            Dim slide As ISlide = presentation.Slides(0)

            ' Define columns with widths and rows with heights
            Dim dblCols As Double() = {120, 120, 120, 120}
            Dim dblRows As Double() = {100, 100, 100, 100}

            ' Add table shape to slide
            Dim tbl As ITable = slide.Shapes.AddTable(100, 50, dblCols, dblRows)
            tbl(1, 0).TextFrame.Text = "10"
            tbl(2, 0).TextFrame.Text = "20"
            tbl(3, 0).TextFrame.Text = "30"

            ' Accessing the text frame
            Dim txtFrame As ITextFrame = tbl(0, 0).TextFrame

            ' Create the Paragraph object for text frame
            Dim paragraph As IParagraph = txtFrame.Paragraphs(0)

            ' Create Portion object for paragraph
            Dim portion As IPortion = paragraph.Portions(0)
            portion.Text = "Text here"
            portion.PortionFormat.FillFormat.FillType = FillType.Solid
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black

            ' Aligning the text vertically
            Dim cell As ICell = tbl(0, 0)
            cell.TextAnchorType = TextAnchorType.Center
            cell.TextVerticalType = TextVerticalType.Vertical270

            ' Save Presentation
            presentation.Save(dataDir + "Vertical_Align_Text_out.pptx", SaveFormat.Pptx)
        End Sub
        ' ExEnd:VerticallyAlignText
    End Class
End Namespace