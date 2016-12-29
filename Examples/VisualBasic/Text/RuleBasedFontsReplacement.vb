Imports System
Imports System.Drawing
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class RuleBasedFontsReplacement
        Public Shared Sub Run()
            ' ExStart:RuleBasedFontsReplacement
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Load presentation
            Dim presentation As New Presentation(dataDir & Convert.ToString("Fonts.pptx"))

            ' Load source font to be replaced
            Dim sourceFont As IFontData = New FontData("SomeRareFont")

            ' Load the replacing font
            Dim destFont As IFontData = New FontData("Arial")

            ' Add font rule for font replacement
            Dim fontSubstRule As IFontSubstRule = New FontSubstRule(sourceFont, destFont, FontSubstCondition.WhenInaccessible)

            ' Add rule to font substitute rules collection
            Dim fontSubstRuleCollection As IFontSubstRuleCollection = New FontSubstRuleCollection()
            fontSubstRuleCollection.Add(fontSubstRule)

            ' Add font rule collection to rule list
            presentation.FontsManager.FontSubstRuleList = fontSubstRuleCollection

            ' Arial font will be used instead of SomeRareFont when inaccessible
            Dim bmp As Bitmap = presentation.Slides(0).GetThumbnail(1.0F, 1.0F)

            ' Save the image to disk in JPEG format
            bmp.Save(dataDir & Convert.ToString("Thumbnail_out.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg)
            ' ExEnd:RuleBasedFontsReplacement

        End Sub
    End Class
End Namespace
