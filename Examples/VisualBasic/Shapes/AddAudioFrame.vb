Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddAudioFrame
        Public Shared Sub Run()

            'ExStart:AddAudioFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Presentation class that represents the PPTX
            Using presentation As New Presentation()

                ' Get the first slide
                Dim slide As ISlide = presentation.Slides(0)

                ' Load the wav sound file to stram
                Dim fileStream As New FileStream(dataDir & "sampleaudio.wav", FileMode.Open, FileAccess.Read)

                ' Add Audio Frame
                Dim audioFrame As IAudioFrame = slide.Shapes.AddAudioFrameEmbedded(50, 150, 100, 100, fileStream)

                ' Set Play Mode and Volume of the Audio
                audioFrame.PlayMode = AudioPlayModePreset.Auto
                audioFrame.Volume = AudioVolumeMode.Loud

                ' Write the PPTX file to disk
                presentation.Save(dataDir & "AudioFrameEmbed_out.pptx", SaveFormat.Pptx)

            End Using
			'ExEnd:AddAudioFrame
        End Sub
    End Class
End Namespace