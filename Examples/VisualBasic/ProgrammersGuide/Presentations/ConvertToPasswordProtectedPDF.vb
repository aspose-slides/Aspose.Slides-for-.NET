'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
Imports Aspose.Slides
Imports VisualBasic

Namespace ProgrammersGuide.Presentations
    Public Class ConvertToPasswordProtectedPDF
        Public Shared Sub Run()

            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As Presentation = New Presentation(dataDir & "DemoFile.pptx")

                ' Instantiate the PdfOptions class
                Dim opts As Export.PdfOptions = New Export.PdfOptions()

                ' Setting PDF password
                opts.Password = "password"

                ' Save the presentation to password protected PDF
                presentation.Save(dataDir & "PasswordProtectedPDF.pdf", Export.SaveFormat.Pdf, opts)
            End Using
        End Sub
    End Class
End Namespace