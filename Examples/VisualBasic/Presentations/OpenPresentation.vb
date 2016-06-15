Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class OpenPresentation
        Public Shared Sub Run()
            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Opening the presentation file by passing the file path to the constructor of Presentation class
            Dim presentation As New Presentation(dataDir & "OpenPresentation.pptx")

            ' Printing the total number of slides present in the presentation
            System.Console.WriteLine(presentation.Slides.Count.ToString())

        End Sub
    End Class
End Namespace