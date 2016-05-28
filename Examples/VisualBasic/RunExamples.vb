Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports ProgrammersGuide.Presentations

Imports VisualBasic.Charts
Imports VisualBasic.Presentations
Imports VisualBasic.Shapes
Imports VisualBasic.Slides
Imports VisualBasic.SmartArts
Imports VisualBasic.Tables
Imports VisualBasic.Text

Namespace VisualBasic
    Friend Class RunExamples
        <STAThread()> _
        Public Shared Sub Main()
            Console.WriteLine("Open RunExamples.cs. In Main() method, Un-comment the example that you want to run")
            Console.WriteLine("=====================================================")
            ' Un-comment the one you want to try out

            ' =====================================================
            ' =====================================================
            ' Charts
            ' =====================================================
            ' =====================================================

            'ChartEntities.Run()
            'ChartTrendLines.Run()
            'ExistingChart.Run()
            'NormalCharts.Run()
            'NumberFormat.Run()
            'ScatteredChart.Run()


            '// =====================================================
            '// =====================================================
            '// Presentations
            '// =====================================================
            '// =====================================================

            AccessBuiltinProperties.Run()
            'AccessModifyingProperties.Run()
            'AccessOpenDoc.Run()
            'AccessProperties.Run()
            'ConvertToPDF.Run()
            'ConvertPDFwithCustomOptions.Run()
            'ConvertToPasswordProtectedPDF.Run()
            'ConvertSpecificSlideToPDF.Run()
            'ConvertSlidesToPdfNotes.Run()
            'ConvertWithNoteToTiff.Run()
            'Convert_HTML.Run()
            'ConvertIndividualSlide.Run()
            'Convert_Tiff_Custom.Run()
            'Convert_Tiff_Default.Run()
            'ConvertToTIFFWithCustomImagePixelFormat.Run()
            'Convert_XPS.Run()
            'Convert_XPS_Options.Run()
            'ModifyBuiltinProperties.Run()
            'OpenPasswordPresentation.Run()
            'VerifyingPresentationWithoutloading.Run()
            'OpenPresentation.Run()
            'PPTtoPPTX.Run()
            'RemoveWriteProtection.Run()
            'SaveAsReadOnly.Run()
            'SaveProperties.Run()
            'SaveToFile.Run()
            'SaveToStream.Run()
            'SaveWithPassword.Run()
            'SaveAsPredefinedViewType.Run()
            'GetFileFormat.Run()
            'ExportMediaFilestohtml.Run()
            'AddCustomDocumentProperties.Run()
            'ConvetToSWF.Run()
            'ConversionToTIFFNotes.Run()
            'ConvertNotesSlideViewToPDF.Run()

            '// =====================================================
            '// =====================================================
            '// Shapes
            '// =====================================================
            '// =====================================================

            'AccessOLEObjectFrame.Run();
            'AddArrowShapedLine.Run();
            'AddArrowShapedLineToSlide.Run();
            'AddAudioFrame.Run();
            'AddOLEObjectFrame.Run();
            'AddPlainLineToSlide.Run();
            'AddSimplePictureFrames.Run();
            'AddVideoFrame.Run();
            'AnimationsOnShapes.Run();
            'ChangeOLEObjectData.Run();
            'ConnectorLineAngle.Run();
            'EmbeddedVideoFrame.Run();
            'FillShapesGradient.Run();
            'FillShapesPattern.Run();
            'FillShapesPicture.Run();
            'FindShapeInSlide.Run();
            'FormatJoinStyles.Run();
            'FormatLines.Run();
            'FormattedEllipse.Run();
            'FormattedRectangle.Run();
            'PictureFrameFormatting.Run();
            'RotatingShapes.Run();
            'SimpleEllipse.Run();
            'SimpleRectangle.Run();


            '// =====================================================
            '// =====================================================
            '// Slides in Presentation
            '// =====================================================
            '// =====================================================

            'AccessSlides.Run();
            'AddSlides.Run();
            'BetterSlideTransitions.Run();
            'ChangePosition.Run();
            'CloneAtEndOfAnother.Run();
            'CloneAtEndOfAnotherSpecificPosition.Run();
            'CloneToAnotherPresentationWithMaster.Run();
            'CloneWithInSamePresentation.Run();
            'CloneWithinSamePresentationToEnd.Run();
            'CreateSlidesSVGImage.Run();
            'RemoveSlideUsingIndex.Run();
            'RemoveSlideUsingReference.Run();
            'SetBackgroundToGradient.Run();
            'SetImageAsBackground.Run();
            'SetSlideBackgroundMaster.Run();
            'SetSlideBackgroundNormal.Run();
            'SimpleSlideTransitions.Run();
            'ThumbnailFromSlide.Run();
            'ThumbnailFromSlideInNotes.Run();
            'ThumbnailWithUserDefinedDimensions.Run();

            '// =====================================================
            '// =====================================================
            '// Smart Arts
            '// =====================================================
            '// =====================================================

            'AccessChildNodes.Run();
            'AccessChildNodeSpecificPosition.Run();
            'AccessSmartArt.Run();
            'AccessSmartArtShape.Run();
            'AddNodes.Run();
            'AddNodesSpecificPosition.Run();
            'AssistantNode.Run();
            'CreateSmartArtShape.Run();
            'RemoveNode.Run();
            'RemoveNodeSpecificPosition.Run();
            'SmartArtNodeLevel.Run();

            '// =====================================================
            '// =====================================================
            '// Tables
            '// =====================================================
            '// =====================================================

            'RemovingRowColumn.Run();
            'TableFromScratch.Run();
            'TableWithCellBorders.Run();
            'UpdateExistingTable.Run();

            '// =====================================================
            '// =====================================================
            '// Text
            '// =====================================================
            '// =====================================================

            'DefaultFonts.Run();
            'ExportingHTMLText.Run();
            'FontFamily.Run();
            'FontProperties.Run();
            'ImportingHTMLText.Run();
            'MultipleParagraphs.Run();
            'ParagraphBullets.Run();
            'ParagraphIndent.Run();
            'ParagraphsAlignment.Run();
            'ReplacingText.Run();
            'ShadowEffects.Run();
            'TextBoxHyperlink.Run();
            'TextBoxOnSlideProgram.Run();


            ' Stop before exiting
            Console.WriteLine(Constants.vbLf + Constants.vbLf & "Program Finished. Press any key to exit....")
            Console.ReadKey()
        End Sub

        Public Shared Function GetDataDir_Charts() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Charts/Data/")
        End Function

        Public Shared Function GetDataDir_Presentations() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Presentations/Data/")
        End Function

        Public Shared Function GetDataDir_Shapes() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Shapes/Data/")
        End Function

        Public Shared Function GetDataDir_Slides_Presentations() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Slides-Presentations/Data/")
        End Function

        Public Shared Function GetDataDir_SmartArts() As String
            Return Path.GetFullPath("../../ProgrammersGuide/SmartArts/Data/")
        End Function

        Public Shared Function GetDataDir_Tables() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Tables/Data/")
        End Function

        Public Shared Function GetDataDir_Text() As String
            Return Path.GetFullPath("../../ProgrammersGuide/Text/Data/")
        End Function
    End Class
End Namespace