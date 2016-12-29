Imports System
Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class EmbeddedVideoFrame
        Public Shared Sub Run()
			'ExStart:EmbeddedVideoFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()
            Dim videoDir As String = RunExamples.GetDataDir_Video()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If Not IsExists Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If
            ' Instantiate Presentation class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Embedd vide inside presentation
                Dim vid As IVideo = pres.Videos.AddVideo(New FileStream(videoDir & Convert.ToString("Wildlife.mp4"), FileMode.Open))

                ' Add Video Frame
                Dim vf As IVideoFrame = sld.Shapes.AddVideoFrame(50, 150, 300, 350, vid)

                ' Set video to Video Frame
                vf.EmbeddedVideo = vid

                ' Set Play Mode and Volume of the Video
                vf.PlayMode = VideoPlayModePreset.Auto
                vf.Volume = AudioVolumeMode.Loud

                'Write the PPTX file to disk
                pres.Save(dataDir & Convert.ToString("VideoFrame_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:EmbeddedVideoFrame
        End Sub
    End Class
End Namespace