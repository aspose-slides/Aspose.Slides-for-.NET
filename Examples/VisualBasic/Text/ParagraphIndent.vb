Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class ParagraphIndent
        Public Shared Sub Run()
            ' ExStart:ParagraphIndent
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Presentation Class
            Dim pres As New Presentation()

            ' Get first slide
            Dim sld As ISlide = pres.Slides(0)

            ' Add a Rectangle Shape
            Dim rect As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 500, 150)

            ' Add TextFrame to the Rectangle
            Dim tf As ITextFrame = rect.AddTextFrame("This is first line " & Constants.vbCr & "This is second line " & Constants.vbCr & "This is third line")

            ' Set the text to fit the shape
            tf.TextFrameFormat.AutofitType = TextAutofitType.Shape

            ' Hide the lines of the Rectangle
            rect.LineFormat.FillFormat.FillType = FillType.Solid

            ' Get first Paragraph in the TextFrame and set its Indent
            Dim para1 As IParagraph = tf.Paragraphs(0)
            ' Setting paragraph bullet style and symbol
            para1.ParagraphFormat.Bullet.Type = BulletType.Symbol
            para1.ParagraphFormat.Bullet.Char = Convert.ToChar(8226)
            para1.ParagraphFormat.Alignment = TextAlignment.Left

            para1.ParagraphFormat.Depth = 2
            para1.ParagraphFormat.Indent = 30

            ' Get second Paragraph in the TextFrame and set its Indent
            Dim para2 As IParagraph = tf.Paragraphs(1)
            para2.ParagraphFormat.Bullet.Type = BulletType.Symbol
            para2.ParagraphFormat.Bullet.Char = Convert.ToChar(8226)
            para2.ParagraphFormat.Alignment = TextAlignment.Left
            para2.ParagraphFormat.Depth = 2
            para2.ParagraphFormat.Indent = 40

            ' Get third Paragraph in the TextFrame and set its Indent
            Dim para3 As IParagraph = tf.Paragraphs(2)
            para3.ParagraphFormat.Bullet.Type = BulletType.Symbol
            para3.ParagraphFormat.Bullet.Char = Convert.ToChar(8226)
            para3.ParagraphFormat.Alignment = TextAlignment.Left
            para3.ParagraphFormat.Depth = 2
            para3.ParagraphFormat.Indent = 50

            'Write the Presentation to disk
            pres.Save(dataDir & "InOutDent_out.pptx", SaveFormat.Pptx)
            ' ExEnd:ParagraphIndent
        End Sub
    End Class
End Namespace