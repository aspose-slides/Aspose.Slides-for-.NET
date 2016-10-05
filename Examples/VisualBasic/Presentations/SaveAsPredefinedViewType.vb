Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class SaveAsPredefinedViewType
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Opening the presentation file
            Dim presentation As Presentation = New Presentation()

            ' Setting view type
            presentation.ViewProperties.LastView = ViewType.SlideMasterView

            ' Saving presentation
            presentation.Save(dataDir & "SetViewType_out.pptx", Export.SaveFormat.Pptx)

        End Sub
    End Class
End Namespace