Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class CreateSlidesSVGImage
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate a Presentation class that represents the presentation file

            Using pres As New Presentation(dataDir & "CreateSlidesSVGImage.pptx")

                ' Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Create a memory stream object
                Dim SvgStream As New MemoryStream()

                ' Generate SVG image of slide and save in memory stream
                sld.WriteAsSvg(SvgStream)
                SvgStream.Position = 0

                ' Save memory stream to file
                Using fileStream As Stream = System.IO.File.OpenWrite(dataDir & "Aspose5_out.svg")
                    Dim buffer(8 * 1024 - 1) As Byte
                    Dim len As Integer
                    len = SvgStream.Read(buffer, 0, buffer.Length)
                    Do While len > 0
                        fileStream.Write(buffer, 0, len)
                        len = SvgStream.Read(buffer, 0, buffer.Length)
                    Loop

                End Using
                SvgStream.Close()
            End Using
        End Sub
    End Class
End Namespace