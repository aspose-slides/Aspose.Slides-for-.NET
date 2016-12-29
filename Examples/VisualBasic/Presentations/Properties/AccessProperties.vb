Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class AccessProperties
        Public Shared Sub Run()
			'ExStart:AccessProperties
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationProperties()

            ' Accessing the Document Properties of a Password Protected Presentation without Password
            ' creating instance of load options to set the presentation access password
            Dim loadOptions As New LoadOptions()

            ' Setting the access password to null
            loadOptions.Password = Nothing

            ' Setting the access to document properties
            loadOptions.OnlyLoadDocumentProperties = True

            ' Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Dim pres As New Presentation(dataDir & Convert.ToString("AccessProperties.pptx"), loadOptions)

            ' Getting Document Properties
            Dim docProps As IDocumentProperties = pres.DocumentProperties

            System.Console.WriteLine("Name of Application : " + docProps.NameOfApplication)
			'ExEnd:AccessProperties
        End Sub
    End Class
End Namespace
