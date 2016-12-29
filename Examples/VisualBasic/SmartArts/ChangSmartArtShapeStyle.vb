Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.SmartArt
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Class ChangSmartArtShapeStyle
        Public Shared Sub Run()
            ' ExStart:ChangSmartArtShapeStyle
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            Using presentation As New Presentation(dataDir & Convert.ToString("AccessSmartArtShape.pptx"))
                ' Traverse through every shape inside first slide
                For Each shape As IShape In presentation.Slides(0).Shapes
                    ' Check if shape is of SmartArt type
                    Dim smart As ISmartArt = TryCast(shape, ISmartArt)
                    If (smart IsNot Nothing) Then
                        ' Typecast shape to SmartArtEx
                        ' Checking SmartArt style
                        If smart.QuickStyle = SmartArtQuickStyleType.SimpleFill Then
                            ' Changing SmartArt Style
                            smart.QuickStyle = SmartArtQuickStyleType.Cartoon
                        End If
                    End If
                Next
                ' Saving Presentation
                presentation.Save(dataDir & Convert.ToString("ChangeSmartArtStyle_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:ChangSmartArtShapeStyle
        End Sub
    End Class
End Namespace