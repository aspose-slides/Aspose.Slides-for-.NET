Imports System.Drawing
Imports System.Drawing.Imaging
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace VisualBasic.Shapes
    Public Class CreateShapeThumbnail
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate a Presentation class that represents the presentation file
            Using p As New Presentation(dataDir + "HelloWorld.pptx")

                ' Create a full scale image
                Using bitmap As Bitmap = p.Slides(0).Shapes(0).GetThumbnail()

                    ' Save the image to disk in PNG format
                    bitmap.Save(dataDir + "Shape_thumbnail.png", ImageFormat.Png)

                End Using

            End Using

        End Sub
    End Class
End Namespace