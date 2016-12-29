Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Opening
    Public Class OpenPasswordPresentation
        Public Shared Sub Run()
			'ExStart:OpenPasswordPresentation
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationOpening()

            ' creating instance of load options to set the presentation access password
            Dim loadOptions As New LoadOptions()

            ' Setting the access password
            loadOptions.Password = "pass"

            ' Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Dim pres As New Presentation(dataDir & Convert.ToString("OpenPasswordPresentation.pptx"), loadOptions)

            ' Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString())
			'ExEnd:OpenPasswordPresentation
		End Sub
    End Class
End Namespace
