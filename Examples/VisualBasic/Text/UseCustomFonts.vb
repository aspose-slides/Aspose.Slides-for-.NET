Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.Charts
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace VisualBasic.Text
    Class UseCustomFonts
        Public Shared Sub Run()

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            Dim loadFonts As [String]() = New [String]() {dataDir & Convert.ToString("CustomFonts.ttf")}

            ' Load the custom font directory fonts
            FontsLoader.LoadExternalFonts(loadFonts)

            ' Do Some work and perform presentation/slides rendering
            Using presentation As New Presentation(dataDir & Convert.ToString("DefaultFonts.pptx"))
                presentation.Save(dataDir & Convert.ToString("NewFonts.pptx"), SaveFormat.Pptx)
            End Using

            ' Clear Font Cachce
            FontsLoader.ClearCache()

        End Sub
    End Class
End Namespace
