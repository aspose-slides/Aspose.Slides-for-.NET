Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.VBA
    Class RemoveVBAMacros
        Public Shared Sub Run()
            ' ExStart:RemoveVBAMacros
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_VBA()
            ' Instantiate Presentation
            Using presentation As New Presentation(dataDir & Convert.ToString("VBA.pptm"))
                ' Access the Vba module and remove 
                presentation.VbaProject.Modules.Remove(presentation.VbaProject.Modules(0))

                ' Save Presentation
                presentation.Save(dataDir & Convert.ToString("RemovedVBAMacros_out.pptm"), SaveFormat.Pptm)
            End Using
            ' ExEnd:RemoveVBAMacros
        End Sub
    End Class
End Namespace