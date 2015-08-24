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
    Public Class AddAudioFrame
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Load the wav sound file to stram
                Dim fstr As New FileStream(dataDir & "sampleaudio.wav", FileMode.Open, FileAccess.Read)

                'Add Audio Frame
                Dim af As IAudioFrame = sld.Shapes.AddAudioFrameEmbedded(50, 150, 100, 100, fstr)

                'Set Play Mode and Volume of the Audio
                af.PlayMode = AudioPlayModePreset.Auto
                af.Volume = AudioVolumeMode.Loud

                'Write the PPTX file to disk
                pres.Save(dataDir & "AudioFrameEmbed.pptx", SaveFormat.Pptx)
            End Using



        End Sub
    End Class
End Namespace