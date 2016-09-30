Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Aspose.Slides.Examples.VisualBasic.ActiveX
Imports Aspose.Slides.Examples.VisualBasic.Presentations
Imports Aspose.Slides.Examples.VisualBasic.Charts
Imports Aspose.Slides.Examples.VisualBasic.Rendering.Printing
Imports Aspose.Slides.Examples.VisualBasic.Shapes
Imports Aspose.Slides.Examples.VisualBasic.Slides
Imports Aspose.Slides.Examples.VisualBasic.SmartArts
Imports Aspose.Slides.Examples.VisualBasic.Tables
Imports Aspose.Slides.Examples.VisualBasic.Text
Imports Aspose.Slides.Examples.VisualBasic.VBA

Namespace Aspose.Slides.Examples.VisualBasic
    Friend Class RunExamples
        <STAThread()> _
        Public Shared Sub Main()
            Console.WriteLine("Open RunExamples.vb." & vbLf & "In Main() method, Un-comment the example that you want to run")
            Console.WriteLine("=====================================================")
            ' Un-comment the one you want to try out


            '// =====================================================
            '// =====================================================
            '//  Active X
            '// =====================================================
            '// =====================================================

            'ManageActiveXControl.Run()
            'LinkingVideoActiveXControl.Run()

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
            'PieChart.Run()
            'ChangeChartCategoryAxis.Run()
            'DisplayChartLabels.Run()
            'AddErrorBars.Run()
            'AddCustomError.Run()
            'AnimatingSeries.Run()
            'AnimatingSeriesElements.Run()
            'AnimatingCategoriesElements.Run()
            'SetChartSeriesOverlap.Run()
            'SetAutomaticSeriesFillColor.Run()
            'SetCategoryAxisLabelDistance.Run()
            'SetlegendCustomOptions.Run()
            'SetDataLabelsPercentageSign.Run()
            'DoughnutChartHole.Run()
            'ManagePropertiesCharts.Run()
            'SetGapWidth.Run()
            'AutomaticChartSeriescolor.Run()
            'DisplayPercentageAsLabels.Run()
            'SecondPlotOptionsforCharts.Run()
            'SetMarkerOptions.Run()

            '// =====================================================
            '// =====================================================
            '// Presentations
            '// =====================================================
            '// =====================================================

            'AccessBuiltinProperties.Run()
            'AccessModifyingProperties.Run()
            'AccessOpenDoc.Run()
            'AccessProperties.Run()
            'ConvertToPDF.Run()
            'ConvertPDFwithCustomOptions.Run()
            'ConvertToPDFWithHiddenSlides.Run()
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
            'ConversionToTIFFNotes.Run()
            'ConvertNotesSlideViewToPDF.Run()
            'CreateNewPresentation.Run()
            'ConvetToSWF.Run()
            'GetRectangularCoordinatesofParagraph.Run()
            'GetPositionCoordinatesofPortion.Run()

            '// =====================================================
            '// =====================================================
            '// Shapes
            '// =====================================================
            '// =====================================================

            'AccessOLEObjectFrame.Run()
            'AddArrowShapedLine.Run()
            'AddArrowShapedLineToSlide.Run()
            'AddAudioFrame.Run()
            'AddOLEObjectFrame.Run()
            'AddPlainLineToSlide.Run()
            'AddSimplePictureFrames.Run()
            'AddVideoFrame.Run()
            'AnimationsOnShapes.Run()
            'ChangeOLEObjectData.Run()
            'ConnectorLineAngle.Run()
            'EmbeddedVideoFrame.Run()
            'FillShapesGradient.Run()
            'FillShapesPattern.Run()
            'FillShapesPicture.Run()
            'FillShapeswithSolidColor.Run()
            'FindShapeInSlide.Run()
            'FormatJoinStyles.Run()
            'FormatLines.Run()
            'FormattedEllipse.Run()
            'FormattedRectangle.Run()
            'PictureFrameFormatting.Run()
            'RotatingShapes.Run()
            'SimpleEllipse.Run()
            'SimpleRectangle.Run()
            'AddRelativeScaleHeightPictureFrame.Run()
            'CreateShapeThumbnail.Run()
            'CreateBoundsShapeThumbnail.Run()
            'CreateSmartArtChildNoteThumbnail.Run()
            'CreateScalingFactorThumbnail.Run()
            'CreateGroupShape.Run()
            'AccessingAltTextinGroupshapes.Run()
            'CloneShapes.Run()
            'SettAlternativeText.Run()
            'RemoveShape.Run()
            'Hidingshapes.Run()
            'ChangeShapeOrder.Run()
            'ConnectShapesUsingConnectors.Run()
            'ConnectShapeUsingConnectionSite.Run()
            'ApplyBevelEffects.Run()

            '// =====================================================
            '// =====================================================
            '// Slides  
            '// =====================================================
            '// =====================================================

            'AccessSlides.Run()
            'AddSlides.Run()
            'BetterSlideTransitions.Run()
            'ChangePosition.Run()
            'CloneAtEndOfAnother.Run()
            'CloneAtEndOfAnotherSpecificPosition.Run()
            'CloneToAnotherPresentationWithMaster.Run()
            'CloneWithInSamePresentation.Run()
            'CloneWithinSamePresentationToEnd.Run()
            'CreateSlidesSVGImage.Run()
            'RemoveSlideUsingIndex.Run()
            'RemoveSlideUsingReference.Run()
            'SetBackgroundToGradient.Run()
            'SetImageAsBackground.Run()
            'SetSlideBackgroundMaster.Run()
            'SetSlideBackgroundNormal.Run()
            'SimpleSlideTransitions.Run()
            'ThumbnailFromSlide.Run()
            'ThumbnailFromSlideInNotes.Run()
            'ThumbnailWithUserDefinedDimensions.Run()
            'AccessSlidebyIndex.Run()
            'AccessSlidebyID.Run()
            'CloneAnotherPresentationAtSpecifiedPosition.Run()
            'ManagSimpleSlideTransitions.Run()
            'ManagingBetterSlideTransitions.Run()
            'AddSlideComments.Run()
            'AccessSlideComments.Run()
            'RemoveHyperlinks.Run()
            'AddLayoutSlides.Run()
            'SettSizeAndType.Run()
            'SetPDFPageSize.Run()
            'RemoveNotesAtSpecificSlide.Run()
            'RemoveNotesFromAllSlides.Run()
            'ExtractVideo.Run()
            'SetTransitionEffects.Run()

            '// =====================================================
            '// =====================================================
            '// Rendering
            '// =====================================================
            '// =====================================================

            'SetZoom.Run()
            'SetSlideNumber.Run()
            'DefaultPrinterPrinting.Run()
            'SpecificPrinterPrinting.Run()

            '// =====================================================
            '// =====================================================
            '// Smart Arts
            '// =====================================================
            '// =====================================================

            'AccessChildNodes.Run()
            'AccessChildNodeSpecificPosition.Run()
            'AccessSmartArt.Run()
            'AccessSmartArtShape.Run()
            'AddNodes.Run()
            'AddNodesSpecificPosition.Run()
            'AssistantNode.Run()
            'CreateSmartArtShape.Run()
            'RemoveNode.Run()
            'RemoveNodeSpecificPosition.Run()
            'SmartArtNodeLevel.Run()
            'AccessSmartArtParticularLayout.Run()
            'ChangSmartArtShapeStyle.Run()
            'ChangeSmartArtShapeColorStyle.Run()
            'FillFormatSmartArtShapeNode.Run()
            'ChangeTextOnSmartArtNode.Run()
            'ChangeSmartArtLayout.Run()
            'CheckSmartArtHiddenProperty.Run()
            'ChangeSmartArtState.Run()
            'OrganizeChartLayoutType.Run()

            '// =====================================================
            '// =====================================================
            '// Tables
            '// =====================================================
            '// =====================================================

            'RemovingRowColumn.Run()
            'TableFromScratch.Run()
            'TableWithCellBorders.Run()
            'UpdateExistingTable.Run()
            'AddImageinsideTableCell.Run()
            'CloningInTable.Run()
            'VerticallyAlignText.Run()
            'StandardTables.Run()
            'MergeCells.Run()
            'MergeCell.Run()
            'CellSplit.Run()


            '// =====================================================
            '// =====================================================
            '// Text
            '// =====================================================
            '// =====================================================

            'DefaultFonts.Run()
            'ExportingHTMLText.Run()
            'FontFamily.Run()
            'FontProperties.Run()
            'ImportingHTMLText.Run()
            'MultipleParagraphs.Run()
            'ParagraphBullets.Run()
            'ParagraphIndent.Run()
            'ParagraphsAlignment.Run()
            'ReplacingText.Run()
            'ShadowEffects.Run()
            'TextBoxHyperlink.Run()
            'TextBoxOnSlideProgram.Run()
            'ApplyInnerShadow.Run()
            'ManagParagraphFontProperties.Run()
            'SetTextFontProperties.Run()
            'ReplaceFontsExplicitly.Run()
            'RuleBasedFontsReplacement.Run()
            'SetAutofitOftextframe.Run()
            'SetAnchorOfTextFrame.Run()
            'RotatingText.Run()
            'LineSpacing.Run()
            'ApplyOuterShadow.Run()
            'CustomRotationAngleTextframe.Run()
            'UseCustomFonts.Run()
            'ManageParagraphPictureBulletsInPPT.Run()

            '// =====================================================
            '// =====================================================
            '// Working With VBA
            '// =====================================================
            '// =====================================================

            'AddVBAMacros.Run()
            'RemoveVBAMacros.Run()


            ' Stop before exiting
            Console.WriteLine(Constants.vbLf + Constants.vbLf & "Program Finished. Press any key to exit....")
            Console.ReadKey()
        End Sub


        Public Shared Function GetDataDir_ActiveX() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("ActiveX/"))
        End Function
        Public Shared Function GetDataDir_Charts() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Charts/"))
        End Function
        Public Shared Function GetDataDir_VBA() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("VBA/"))
        End Function
        Public Shared Function GetDataDir_Presentations() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Presentations/"))
        End Function

        Public Shared Function GetDataDir_Rendering() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Rendering-Printing/"))
        End Function

        Public Shared Function GetDataDir_Shapes() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Shapes/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/"))
        End Function

        Public Shared Function GetDataDir_SmartArts() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("SmartArts/"))
        End Function

        Public Shared Function GetDataDir_Tables() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Tables/"))
        End Function

        Public Shared Function GetDataDir_Text() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Text/"))
        End Function

        Public Shared Function GetDataDir_Video() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Video/"))
        End Function

        Private Shared Function GetDataDir_Data() As String
            Dim parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent
            Dim startDirectory As String = Nothing
            If parent IsNot Nothing Then
                Dim directoryInfo = parent.Parent
                If directoryInfo IsNot Nothing Then
                    startDirectory = directoryInfo.FullName
                End If
            Else
                startDirectory = parent.FullName
            End If
            Return If(startDirectory IsNot Nothing, Path.Combine(startDirectory, "Data\"), Nothing)
        End Function

    End Class
End Namespace