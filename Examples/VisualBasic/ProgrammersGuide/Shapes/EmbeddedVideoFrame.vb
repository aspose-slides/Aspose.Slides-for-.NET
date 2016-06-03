'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class EmbeddedVideoFrame
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()


            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If
            'Instantiate Presentation class that represents the PPTX
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Embedd vide inside presentation
                Dim vid As IVideo = pres.Videos.AddVideo(New FileStream(dataDir & "Wildlife.mp4", FileMode.Open))

                'Add Video Frame
                Dim vf As IVideoFrame = sld.Shapes.AddVideoFrame(50, 150, 300, 350, vid)

                'Set video to Video Frame
                vf.EmbeddedVideo = vid

                'Set Play Mode and Volume of the Video
                vf.PlayMode = VideoPlayModePreset.Auto
                vf.Volume = AudioVolumeMode.Loud

                'Write the PPTX file to disk
                pres.Save(dataDir & "VideoFrame.pptx", SaveFormat.Pptx)
            End Using



        End Sub
    End Class
End Namespace