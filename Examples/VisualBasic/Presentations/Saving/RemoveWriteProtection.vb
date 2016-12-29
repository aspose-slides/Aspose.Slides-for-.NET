Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Presentations.Saving
    Public Class RemoveWriteProtection
        Public Shared Sub Run()
			'ExStart:RemoveWriteProtection
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_PresentationSaving()

            ' Opening the presentation file
            Dim presentation As New Presentation(dataDir & Convert.ToString("RemoveWriteProtection.pptx"))

            ' Checking if presentation is write protected
            If presentation.ProtectionManager.IsWriteProtected Then
                ' Removing Write protection                
                presentation.ProtectionManager.RemoveWriteProtection()
            End If

            ' Saving presentation
            presentation.Save(dataDir & Convert.ToString("File_Without_WriteProtection_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
			'ExEnd:RemoveWriteProtection
        End Sub
    End Class
End Namespace
