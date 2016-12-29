Imports System
Imports System.Drawing
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Class FillFormatSmartArtShapeNode
        Public Shared Sub Run()
            ' ExStart:FillFormatSmartArtShapeNode
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            Using presentation As New Presentation()
                ' Accessing the slide
                Dim slide As ISlide = presentation.Slides(0)

                ' Adding SmartArt shape and nodes
                Dim chevron = slide.Shapes.AddSmartArt(10, 10, 800, 60, SmartArtLayoutType.ClosedChevronProcess)
                Dim node = chevron.AllNodes.AddNode()
                node.TextFrame.Text = "Some text"

                ' Setting node fill color
                For Each item As IShape In node.Shapes
                    item.FillFormat.FillType = FillType.Solid
                    item.FillFormat.SolidFillColor.Color = Color.Red
                Next

                ' Saving Presentation
                presentation.Save(dataDir & Convert.ToString("FillFormat_SmartArt_ShapeNode_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:FillFormatSmartArtShapeNode
        End Sub
    End Class
End Namespace