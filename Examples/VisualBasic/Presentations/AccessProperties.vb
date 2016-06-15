Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class AccessProperties
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Accessing the Document Properties of a Password Protected Presentation without Password
            'creating instance of load options to set the presentation access password
            Dim loadOptions As New Aspose.Slides.LoadOptions()

            'Setting the access password to null
            loadOptions.Password = "Password"

            'Setting the access to document properties
            loadOptions.OnlyLoadDocumentProperties = True

            'Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Dim pres As New Presentation(dataDir & "AccessProperties.pptx", loadOptions)

            'Getting Document Properties
            Dim docProps As IDocumentProperties = pres.DocumentProperties

            System.Console.WriteLine("Name of Application : " & docProps.NameOfApplication)
        End Sub
    End Class
End Namespace