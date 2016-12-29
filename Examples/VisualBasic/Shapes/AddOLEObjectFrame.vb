
Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddOLEObjectFrame
        Public Shared Sub Run()
			'ExStart:AddOLEObjectFrame
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Prseetation class that represents the PPTX
            Dim pres As New Presentation()

            ' Access the first slide
            Dim sld As ISlide = pres.Slides(0)

            ' Load an cel file to stream
            Dim fs As New FileStream(dataDir & "book1.xlsx", FileMode.Open, FileAccess.Read)
            Dim mstream As New MemoryStream()
            Dim buf(4095) As Byte

            Do
                Dim bytesRead As Integer = fs.Read(buf, 0, buf.Length)
                If bytesRead <= 0 Then
                    Exit Do
                End If
                mstream.Write(buf, 0, bytesRead)
            Loop

            ' Add an Ole Object Frame shape
            Dim oof As IOleObjectFrame = sld.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height, "Excel.Sheet.12", mstream.ToArray())

            'Write the PPTX to disk
            pres.Save(dataDir & "OleEmbed_out.pptx", SaveFormat.Pptx)
			'ExEnd:AddOLEObjectFrame
        End Sub
    End Class
End Namespace