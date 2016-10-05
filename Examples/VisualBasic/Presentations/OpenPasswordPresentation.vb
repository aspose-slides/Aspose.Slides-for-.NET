Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class OpenPasswordPresentation
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Creating instance of load options to set the presentation access password
            Dim loadOptions As New Aspose.Slides.LoadOptions()

            ' Setting the access password
            loadOptions.Password = "pass"

            ' Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Dim pres As New Presentation(dataDir & "OpenPasswordPresentation.pptx", loadOptions)

            ' Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString())
        End Sub
    End Class
End Namespace