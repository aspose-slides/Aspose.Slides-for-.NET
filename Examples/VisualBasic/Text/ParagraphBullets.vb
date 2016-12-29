Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class ParagraphBullets
        Public Shared Sub Run()
            ' ExStart:ParagraphBullets
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Creating a presenation instance
            Using pres As New Presentation()

                ' Accessing the first slide
                Dim slide As ISlide = pres.Slides(0)


                ' Adding and accessing Autoshape
                Dim aShp As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 400, 200)

                ' Accessing the text frame of created autoshape
                Dim txtFrm As ITextFrame = aShp.TextFrame

                ' Removing the default exisiting paragraph
                txtFrm.Paragraphs.RemoveAt(0)

                ' Creating a paragraph
                Dim para As New Paragraph()

                ' Setting paragraph bullet style and symbol
                para.ParagraphFormat.Bullet.Type = BulletType.Symbol
                para.ParagraphFormat.Bullet.Char = Convert.ToChar(8226)

                ' Setting paragraph text
                para.Text = "Welcome to Aspose.Slides"

                ' Setting bullet indent
                para.ParagraphFormat.Indent = 25

                ' Setting bullet color
                para.ParagraphFormat.Bullet.Color.ColorType = ColorType.RGB
                para.ParagraphFormat.Bullet.Color.Color = Color.Black
                para.ParagraphFormat.Bullet.IsBulletHardColor = NullableBool.True ' set IsBulletHardColor to true to use own bullet color

                ' Setting Bullet Height
                para.ParagraphFormat.Bullet.Height = 100

                ' Adding Paragraph to text frame
                txtFrm.Paragraphs.Add(para)

                ' Creating second paragraph
                Dim para2 As New Paragraph()

                ' Setting paragraph bullet type and style
                para2.ParagraphFormat.Bullet.Type = BulletType.Numbered
                para2.ParagraphFormat.Bullet.NumberedBulletStyle = NumberedBulletStyle.BulletCircleNumWDBlackPlain

                ' Adding paragraph text
                para2.Text = "This is numbered bullet"

                ' Setting bullet indent
                para2.ParagraphFormat.Indent = 25

                para2.ParagraphFormat.Bullet.Color.ColorType = ColorType.RGB
                para2.ParagraphFormat.Bullet.Color.Color = Color.Black
                para2.ParagraphFormat.Bullet.IsBulletHardColor = NullableBool.True ' set IsBulletHardColor to true to use own bullet color

                ' Setting Bullet Height
                para2.ParagraphFormat.Bullet.Height = 100

                ' Adding Paragraph to text frame
                txtFrm.Paragraphs.Add(para2)


                'Writing the presentation as a PPTX file
                pres.Save(dataDir & "Bullet_out.pptx", SaveFormat.Pptx)
            End Using
            ' ExEnd:ParagraphBullets
        End Sub
    End Class
End Namespace