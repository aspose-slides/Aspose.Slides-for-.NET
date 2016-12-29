Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddVideoFrame
        Public Shared Sub Run()
			'ExStart:AddVideoFrame	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate PrseetationEx class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add Video Frame
                Dim vf As IVideoFrame = sld.Shapes.AddVideoFrame(50, 150, 300, 150, dataDir & "video1.avi")

                ' Set Play Mode and Volume of the Video
                vf.PlayMode = VideoPlayModePreset.Auto
                vf.Volume = AudioVolumeMode.Loud

                'Write the PPTX file to disk
                pres.Save(dataDir & "VideoFrame1_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:AddVideoFrame	
        End Sub
    End Class
End Namespace