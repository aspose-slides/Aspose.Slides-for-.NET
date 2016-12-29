Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class MultipleParagraphs
        Public Shared Sub Run()
            ' ExStart:MultipleParagraphs
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a Presentation class that represents a PPTX file
            Using pres As New Presentation()

                ' Accessing first slide
                Dim slide As ISlide = pres.Slides(0)

                ' Add an AutoShape of Rectangle type
                Dim ashp As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 300, 150)

                ' Access TextFrame of the AutoShape
                Dim tf As ITextFrame = ashp.TextFrame

                ' Create Paragraphs and Portions with different text formats
                Dim para0 As IParagraph = tf.Paragraphs(0)
                Dim port01 As IPortion = New Portion()
                Dim port02 As IPortion = New Portion()
                para0.Portions.Add(port01)
                para0.Portions.Add(port02)

                Dim para1 As IParagraph = New Paragraph()
                tf.Paragraphs.Add(para1)
                Dim port10 As IPortion = New Portion()
                Dim port11 As IPortion = New Portion()
                Dim port12 As IPortion = New Portion()
                para1.Portions.Add(port10)
                para1.Portions.Add(port11)
                para1.Portions.Add(port12)

                Dim para2 As IParagraph = New Paragraph()
                tf.Paragraphs.Add(para2)
                Dim port20 As IPortion = New Portion()
                Dim port21 As IPortion = New Portion()
                Dim port22 As IPortion = New Portion()
                para2.Portions.Add(port20)
                para2.Portions.Add(port21)
                para2.Portions.Add(port22)

                For i As Integer = 0 To 2
                    For j As Integer = 0 To 2
                        tf.Paragraphs(i).Portions(j).Text = "Portion0" & j.ToString()
                        If j = 0 Then
                            tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.FillType = FillType.Solid
                            tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.SolidFillColor.Color = Color.Red
                            tf.Paragraphs(i).Portions(j).PortionFormat.FontBold = NullableBool.True
                            tf.Paragraphs(i).Portions(j).PortionFormat.FontHeight = 15
                        ElseIf j = 1 Then
                            tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.FillType = FillType.Solid
                            tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue
                            tf.Paragraphs(i).Portions(j).PortionFormat.FontItalic = NullableBool.True
                            tf.Paragraphs(i).Portions(j).PortionFormat.FontHeight = 18
                        End If
                    Next j
                Next i

                'Write PPTX to Disk
                pres.Save(dataDir & "multiParaPort_out.pptx", SaveFormat.Pptx)
            End Using
            ' ExEnd:MultipleParagraphs
        End Sub
    End Class
End Namespace