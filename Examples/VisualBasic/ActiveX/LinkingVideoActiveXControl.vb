Imports System
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.ActiveX
    Public Class LinkingVideoActiveXControl
        Public Shared Sub Run()
        	'ExStart:LinkingVideoActiveXControl
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_ActiveX()
            Dim dataVideo As String = RunExamples.GetDataDir_Video()

            ' Instantiate Presentation class that represents PPTX file
            Dim presentation As New Presentation(dataDir & Convert.ToString("template.pptx"))

            ' Create empty presentation instance
            Dim newPresentation As New Presentation()

            ' Remove default slide
            newPresentation.Slides.RemoveAt(0)

            ' Clone slide with Media Player ActiveX Control
            newPresentation.Slides.InsertClone(0, presentation.Slides(0))

            ' Access the Media Player ActiveX control and set the video path
            newPresentation.Slides(0).Controls(0).Properties("URL") = dataVideo & Convert.ToString("Wildlife.mp4")

            ' Save the Presentation
            newPresentation.Save(dataDir & Convert.ToString("LinkingVideoActiveXControl_out.pptx"), Aspose.Slides.Export.SaveFormat.Pptx)
            'ExEnd:LinkingVideoActiveXControl
        End Sub
    End Class
End Namespace
