Imports System
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AccessingAltTextinGroupshapes
        Public Shared Sub Run()
			'ExStart:AccessingAltTextinGroupshapes
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()


            ' Instantiate Presentation class that represents PPTX file
            Dim pres As New Presentation(dataDir + "AltText.pptx")

            ' Get the first slide
            Dim sld As ISlide = pres.Slides(0)

            For i As Integer = 0 To sld.Shapes.Count - 1
                ' Accessing the shape collection of slides
                Dim shape As IShape = sld.Shapes(i)

                If TypeOf shape Is GroupShape Then
                    ' Accessing the group shape.
                    Dim grphShape As IGroupShape = DirectCast(shape, IGroupShape)
                    For j As Integer = 0 To grphShape.Shapes.Count - 1
                        Dim shape2 As IShape = grphShape.Shapes(j)
                        ' Accessing the AltText property
                        Console.WriteLine(shape2.AlternativeText)
                    Next
                End If
            Next
			'ExEnd:AccessingAltTextinGroupshapes
        End Sub
    End Class
End Namespace