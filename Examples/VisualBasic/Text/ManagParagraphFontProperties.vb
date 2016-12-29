Imports System
Imports System.Drawing
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class ManagParagraphFontProperties
        Public Shared Sub Run()
            ' ExStart:ManagParagraphFontProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()
            ' Instantiate PresentationEx
            Using presentation As New Presentation(dataDir & Convert.ToString("DefaultFonts.pptx"))
                ' Accessing a slide using its slide position
                Dim slide As ISlide = presentation.Slides(0)

                ' Accessing the first and second placeholder in the slide and typecasting it as AutoShape
                Dim textFrame1 As ITextFrame = DirectCast(slide.Shapes(0), IAutoShape).TextFrame
                Dim textFrame2 As ITextFrame = DirectCast(slide.Shapes(1), IAutoShape).TextFrame

                ' Accessing the first Paragraph
                Dim para1 As IParagraph = textFrame1.Paragraphs(0)
                Dim para2 As IParagraph = textFrame2.Paragraphs(0)

                ' Justify the paragraph
                para2.ParagraphFormat.Alignment = TextAlignment.JustifyLow

                ' Accessing the first portion
                Dim port1 As IPortion = para1.Portions(0)
                Dim port2 As IPortion = para2.Portions(0)

                ' Define new fonts
                Dim fd1 As New FontData("Elephant")
                Dim fd2 As New FontData("Castellar")

                ' Assign new fonts to portion
                port1.PortionFormat.LatinFont = fd1
                port2.PortionFormat.LatinFont = fd2

                ' Set font to Bold
                port1.PortionFormat.FontBold = NullableBool.[True]
                port2.PortionFormat.FontBold = NullableBool.[True]

                ' Set font to Italic
                port1.PortionFormat.FontItalic = NullableBool.[True]
                port2.PortionFormat.FontItalic = NullableBool.[True]

                ' Set font color
                port1.PortionFormat.FillFormat.FillType = FillType.Solid
                port1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Purple
                port2.PortionFormat.FillFormat.FillType = FillType.Solid
                port2.PortionFormat.FillFormat.SolidFillColor.Color = Color.Peru
                ' Write the PPTX to disk 
                presentation.Save(dataDir & Convert.ToString("ManagParagraphFontProperties_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
                ' ExEnd:ManagParagraphFontProperties
            End Using
        End Sub
    End Class
End Namespace
