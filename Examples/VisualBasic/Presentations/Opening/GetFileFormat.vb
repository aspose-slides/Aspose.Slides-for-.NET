Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Opening
    Class GetFileFormat
        Public Shared Sub Run()
			'ExStart:GetFileFormat
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationOpening()

            Dim info As IPresentationInfo = PresentationFactory.Instance.GetPresentationInfo(dataDir & Convert.ToString("HelloWorld.pptx"))

            Select Case info.LoadFormat
                Case LoadFormat.Pptx
                    If True Then
                        Exit Select
                    End If

                Case LoadFormat.Unknown
                    If True Then
                        Exit Select
                    End If
            End Select
			'ExEnd:GetFileFormat
        End Sub
    End Class
End Namespace
