Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides

'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Namespace Aspose.Slides.Examples.VisualBasic.Rendering.Printing
    Public Class SetSlideNumber
        Public Shared Sub Run()
			'ExStart:SetSlideNumber
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Rendering()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("HelloWorld.pptx"))

                ' Get the slide number
                Dim firstSlideNumber As Integer = presentation.FirstSlideNumber

                ' Set the slide number
                presentation.FirstSlideNumber = 10
                presentation.Save(dataDir & Convert.ToString("Set_Slide_Number_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:SetSlideNumber	
        End Sub
    End Class
End Namespace