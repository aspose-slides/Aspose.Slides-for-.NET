Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides.Export
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Shapes
    Public Class CloneShapes
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class 
            Using presentation As Presentation = New Presentation(dataDir + "Source Frame.pptx")

                Dim sourceShapes As IShapeCollection = presentation.Slides(0).Shapes
                Dim blankLayout As ILayoutSlide = presentation.Masters(0).LayoutSlides.GetByType(SlideLayoutType.Blank)

                Dim destSlide As ISlide = presentation.Slides.AddEmptySlide(blankLayout)
                Dim destShapes As IShapeCollection = destSlide.Shapes
                destShapes.AddClone(sourceShapes(1), 50, 150 + sourceShapes(0).Height)
                destShapes.AddClone(sourceShapes(2))
                destShapes.InsertClone(0, sourceShapes(0), 50, 150)

                ' Write the PPTX file to disk
                presentation.Save(dataDir + "CloneShape.pptx", SaveFormat.Pptx)

            End Using

        End Sub
    End Class
End Namespace