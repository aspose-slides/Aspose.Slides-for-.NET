'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports Aspose.Slides

Namespace VisualBasic.Presentations
    Public Class GetFileFormat
        Public Shared Sub Run()
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            Dim info As IPresentationInfo = PresentationFactory.Instance.GetPresentationInfo(dataDir & Convert.ToString("DemoFile.pptx"))

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

        End Sub
    End Class
End Namespace