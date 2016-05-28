'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System.IO
Imports Aspose.Slides
Imports VisualBasic

Namespace ProgrammersGuide.Presentations
    Public Class VerifyingPresentationWithoutloading
        Public Shared Sub Run()

            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Getting the file format using the PresentationFactory class instance

            Dim format As LoadFormat = PresentationFactory.Instance.GetPresentationInfo(dataDir & "DemoFile.pptx").LoadFormat

            'It will return "LoadFormat.Unknown" if the file is other than presentation formats
            
        End Sub
    End Class
End Namespace