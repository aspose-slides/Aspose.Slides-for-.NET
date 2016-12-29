Imports System
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Opening
    Class VerifyingPresentationWithoutloading
        Public Shared Sub Run()
			'ExStart:VerifyingPresentationWithoutloading
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationOpening()

            Dim format As LoadFormat = PresentationFactory.Instance.GetPresentationInfo(dataDir & Convert.ToString("HelloWorld.pptx")).LoadFormat
            ' It will return "LoadFormat.Unknown" if the file is other than presentation formats           
        End Sub
			'ExEnd:VerifyingPresentationWithoutloading
    End Class
End Namespace
