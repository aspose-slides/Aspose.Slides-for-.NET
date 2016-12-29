Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class CreateScalingFactorThumbnail
        Public Shared Sub Run()
			'ExStart:CreateScalingFactorThumbnail
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate a Presentation class that represents the presentation file
            Using presentatio As New Presentation(dataDir & Convert.ToString("HelloWorld.pptx"))

                ' Create a full scale image
                Using bitmap As Bitmap = presentatio.Slides(0).Shapes(0).GetThumbnail(ShapeThumbnailBounds.Shape, 1, 1)

                    ' Save the image to disk in PNG format
                    bitmap.Save(dataDir & Convert.ToString("Scaling Factor Thumbnail_out.png"), ImageFormat.Png)

                End Using

            End Using
			'ExEnd:CreateScalingFactorThumbnail
        End Sub
    End Class
End Namespace