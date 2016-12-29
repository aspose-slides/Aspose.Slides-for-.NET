Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Layout
    Class AddLayoutSlides
        Public Shared Sub Run()
            'ExStart:AddLayoutSlides
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Layout()

            ' Instantiate Presentation class that represents the presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))
                ' Try to search by layout slide type
                Dim layoutSlides As IMasterLayoutSlideCollection = presentation.Masters(0).LayoutSlides
                Dim layoutSlide As ILayoutSlide = If(layoutSlides.GetByType(SlideLayoutType.TitleAndObject), layoutSlides.GetByType(SlideLayoutType.Title))

                If layoutSlide Is Nothing Then
                    ' The situation when a presentation doesn't contain some type of layouts.
                    ' presentation File only contains Blank and Custom layout types.
                    ' But layout slides with Custom types has different slide names,
                    ' like "Title", "Title and Content", etc. And it is possible to use these
                    ' names for layout slide selection.
                    ' Also it is possible to use the set of placeholder shape types. For example,
                    ' Title slide should have only Title pleceholder type, etc.
                    For Each titleAndObjectLayoutSlide As ILayoutSlide In layoutSlides
                        If titleAndObjectLayoutSlide.Name = "Title and Object" Then
                            layoutSlide = titleAndObjectLayoutSlide
                            Exit For
                        End If
                    Next

                    If layoutSlide Is Nothing Then
                        For Each titleLayoutSlide As ILayoutSlide In layoutSlides
                            If titleLayoutSlide.Name = "Title" Then
                                layoutSlide = titleLayoutSlide
                                Exit For
                            End If
                        Next

                        If layoutSlide Is Nothing Then
                            layoutSlide = layoutSlides.GetByType(SlideLayoutType.Blank)
                            If layoutSlide Is Nothing Then
                                layoutSlide = layoutSlides.Add(SlideLayoutType.TitleAndObject, "Title and Object")
                            End If
                        End If
                    End If
                End If

                ' Adding empty slide with added layout slide 
                presentation.Slides.InsertEmptySlide(0, layoutSlide)

                ' Save presentation    
                presentation.Save(dataDir & Convert.ToString("AddLayoutSlides_out.pptx"), SaveFormat.Pptx)
            End Using
            'ExStart:AddLayoutSlides
        End Sub
    End Class
End Namespace