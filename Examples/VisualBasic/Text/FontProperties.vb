Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class FontProperties
        Public Shared Sub Run()
            ' ExStart:FontProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Instantiate a Presentation object that represents a PPTX file//Instantiate a Presentation object that represents a PPTX file
            Using pres As New Presentation(dataDir & "FontProperties.pptx")

                ' Accessing a slide using its slide position
                Dim slide As ISlide = pres.Slides(0)

                ' Accessing the first and second placeholder in the slide and typecasting it as AutoShape
                Dim tf1 As ITextFrame = (CType(slide.Shapes(0), IAutoShape)).TextFrame
                Dim tf2 As ITextFrame = (CType(slide.Shapes(1), IAutoShape)).TextFrame

                ' Accessing the first Paragraph
                Dim para1 As IParagraph = tf1.Paragraphs(0)
                Dim para2 As IParagraph = tf2.Paragraphs(0)

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
                port1.PortionFormat.FontBold = NullableBool.True
                port2.PortionFormat.FontBold = NullableBool.True

                ' Set font to Italic
                port1.PortionFormat.FontItalic = NullableBool.True
                port2.PortionFormat.FontItalic = NullableBool.True

                ' Set font color
                port1.PortionFormat.FillFormat.FillType = FillType.Solid
                port1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Purple
                port2.PortionFormat.FillFormat.FillType = FillType.Solid
                port2.PortionFormat.FillFormat.SolidFillColor.Color = Color.Peru

                'Write the PPTX to disk
                pres.Save(dataDir & "WelcomeFont_out.pptx", SaveFormat.Pptx)
            End Using
            ' ExEnd:FontProperties
        End Sub
    End Class
End Namespace