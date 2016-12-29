Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class TextBoxHyperlink
        Public Shared Sub Run()
            ' ExStart:TextBoxHyperlink

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a Presentation class that represents a PPTX
            Dim pptxPresentation As New Presentation()

            ' Get first slide
            Dim slide As ISlide = pptxPresentation.Slides(0)

            ' Add an AutoShape of Rectangle Type
            Dim pptxShape As IShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 150, 150, 50)

            ' Cast the shape to AutoShape
            Dim pptxAutoShape As IAutoShape = CType(pptxShape, IAutoShape)

            ' Access ITextFrame associated with the AutoShape
            pptxAutoShape.AddTextFrame("")

            Dim ITextFrame As ITextFrame = pptxAutoShape.TextFrame

            ' Add some text to the frame
            ITextFrame.Paragraphs(0).Portions(0).Text = "Aspose.Slides"

            ' Set Hyperlink for the portion text
            Dim HypMan As IHyperlinkManager = ITextFrame.Paragraphs(0).Portions(0).PortionFormat.HyperlinkManager
            HypMan.SetExternalHyperlinkClick("http://www.aspose.com")

            ' Save the PPTX Presentation
            pptxPresentation.Save(dataDir & "hLinkPPTX_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            ' ExEnd:TextBoxHyperlink
        End Sub
    End Class
End Namespace