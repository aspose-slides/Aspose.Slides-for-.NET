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
    Class SetTextFontProperties
        Public Shared Sub Run()
            ' ExStart:SetTextFontProperties

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Instantiate Presentation
            Using presentation As New Presentation()

                ' ExStart:SetTextFontProperties
                ' Get first slide
                Dim sld As ISlide = presentation.Slides(0)

                ' Add an AutoShape of Rectangle type
                Dim ashp As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50)

                ' Remove any fill style associated with the AutoShape
                ashp.FillFormat.FillType = FillType.NoFill

                ' Access the TextFrame associated with the AutoShape
                Dim tf As ITextFrame = ashp.TextFrame
                tf.Text = "Aspose TextBox"

                ' Access the Portion associated with the TextFrame
                Dim port As IPortion = tf.Paragraphs(0).Portions(0)

                ' Set the Font for the Portion
                port.PortionFormat.LatinFont = New FontData("Times New Roman")

                ' Set Bold property of the Font
                port.PortionFormat.FontBold = NullableBool.[True]

                ' Set Italic property of the Font
                port.PortionFormat.FontItalic = NullableBool.[True]

                ' Set Underline property of the Font
                port.PortionFormat.FontUnderline = TextUnderlineType.[Single]

                ' Set the Height of the Font
                port.PortionFormat.FontHeight = 25

                ' Set the color of the Font
                port.PortionFormat.FillFormat.FillType = FillType.Solid
                port.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue

                ' ExEnd:SetTextFontProperties
                ' Write the PPTX to disk 
                presentation.Save(dataDir & Convert.ToString("SetTextFontProperties_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:SetTextFontProperties
        End Sub
    End Class
End Namespace