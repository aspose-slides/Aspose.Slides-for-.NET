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

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class ReplaceFontsExplicitly
        Public Shared Sub Run()
            ' ExStart:ReplaceFontsExplicitly
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Load presentation
            Dim presentation As New Presentation(dataDir & Convert.ToString("Fonts.pptx"))

            ' Load source font to be replaced
            Dim sourceFont As IFontData = New FontData("Arial")

            ' Load the replacing font
            Dim destFont As IFontData = New FontData("Times New Roman")

            ' Replace the fonts
            presentation.FontsManager.ReplaceFont(sourceFont, destFont)

            ' Save the presentation
            presentation.Save(dataDir & Convert.ToString("UpdatedFont_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:ReplaceFontsExplicitly
        End Sub
    End Class
End Namespace