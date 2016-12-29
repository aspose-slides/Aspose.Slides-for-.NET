Imports System
Imports System.Net
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Class AddVideoFrameFromWebSource
        Public Shared Sub Run()
			'ExStart:AddVideoFrameFromWebSource	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            Using pres As New Presentation()
                AddVideoFromYouTube(pres, "Tj75Arhq5ho")
                pres.Save(dataDir & Convert.ToString("AddVideoFrameFromWebSource_out.pptx"), SaveFormat.Pptx)
            End Using
        End Sub

        Private Shared Sub AddVideoFromYouTube(pres As Presentation, videoId As String)
            'add videoFrame
            Dim videoFrame As IVideoFrame = pres.Slides(0).Shapes.AddVideoFrame(10, 10, 427, 240, Convert.ToString("https://www.youtube.com/embed/") & videoId)
            videoFrame.PlayMode = VideoPlayModePreset.Auto

            'load thumbnail
            Using client As New WebClient()
                Dim thumbnailUri As String = (Convert.ToString("http://img.youtube.com/vi/") & videoId) + "/hqdefault.jpg"
                videoFrame.PictureFormat.Picture.Image = pres.Images.AddImage(client.DownloadData(thumbnailUri))
            End Using
			'ExEnd:AddVideoFrameFromWebSource	
        End Sub
    End Class
End Namespace