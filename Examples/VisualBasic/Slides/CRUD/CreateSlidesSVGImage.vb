Imports System
Imports System.IO
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.CRUD
    Public Class CreateSlidesSVGImage
        Public Shared Sub Run()
            ' ExStart:CreateSlidesSVGImage
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_CRUD()

            ' Instantiate a Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("CreateSlidesSVGImage.pptx"))

                ' Access the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Create a memory stream object
                Dim SvgStream As New MemoryStream()

                ' Generate SVG image of slide and save in memory stream
                sld.WriteAsSvg(SvgStream)
                SvgStream.Position = 0

                ' Save memory stream to file
                Using fileStream As Stream = System.IO.File.OpenWrite(dataDir & Convert.ToString("Aspose_out.svg"))
                    Dim buffer As Byte() = New Byte(8 * 1024 - 1) {}
                    Dim len As Integer
                    While (InlineAssignHelper(len, SvgStream.Read(buffer, 0, buffer.Length))) > 0
                        fileStream.Write(buffer, 0, len)

                    End While
                End Using
                SvgStream.Close()
            End Using
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
        ' ExEnd:CreateSlidesSVGImage
    End Class
End Namespace
