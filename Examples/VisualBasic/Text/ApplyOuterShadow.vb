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

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class ApplyOuterShadow
        Public Shared Sub Run()
            ' ExStart:ApplyOuterShadow
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get reference of a slide
            Dim slide As ISlide = presentation.Slides(0)

            ' Add an AutoShape of Rectangle type
            Dim ashp As IAutoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 400, 300)
            ashp.FillFormat.FillType = FillType.NoFill

            ' Add TextFrame to the Rectangle
            ashp.AddTextFrame("Aspose TextBox")
            Dim port As IPortion = ashp.TextFrame.Paragraphs(0).Portions(0)
            Dim pf As IPortionFormat = port.PortionFormat
            pf.FontHeight = 50

            ' Enable InnerShadowEffect    
            Dim ef As IEffectFormat = pf.EffectFormat
            ef.EnableInnerShadowEffect()

            ' Set all necessary parameters
            ef.InnerShadowEffect.BlurRadius = 8.0
            ef.InnerShadowEffect.Direction = 90.0F
            ef.InnerShadowEffect.Distance = 6.0
            ef.InnerShadowEffect.ShadowColor.B = 189

            ' Set ColorType as Scheme
            ef.InnerShadowEffect.ShadowColor.ColorType = ColorType.Scheme

            ' Set Scheme Color
            ef.InnerShadowEffect.ShadowColor.SchemeColor = SchemeColor.Accent1

            ' Save Presentation
            presentation.Save(dataDir & Convert.ToString("WordArt_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:ApplyOuterShadow

        End Sub
    End Class
End Namespace