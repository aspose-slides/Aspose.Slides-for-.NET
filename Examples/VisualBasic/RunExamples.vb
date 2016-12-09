Imports System
Imports System.IO
Imports Aspose.Slides.Examples.VisualBasic.ActiveX
Imports Aspose.Slides.Examples.VisualBasic.Charts
Imports Aspose.Slides.Examples.VisualBasic.Conversion
Imports Aspose.Slides.Examples.VisualBasic.Presentations
Imports Aspose.Slides.Examples.VisualBasic.Presentations.Conversion
Imports Aspose.Slides.Examples.VisualBasic.Presentations.Opening
Imports Aspose.Slides.Examples.VisualBasic.Presentations.Properties
Imports Aspose.Slides.Examples.VisualBasic.Presentations.Saving
Imports Aspose.Slides.Examples.VisualBasic.Rendering.Printing
Imports Aspose.Slides.Examples.VisualBasic.Shapes
Imports Aspose.Slides.Examples.VisualBasic.Slides
Imports Aspose.Slides.Examples.VisualBasic.Slides.Background
Imports Aspose.Slides.Examples.VisualBasic.Slides.Comments
Imports Aspose.Slides.Examples.VisualBasic.Slides.CRUD
Imports Aspose.Slides.Examples.VisualBasic.Slides.Hyperlink
Imports Aspose.Slides.Examples.VisualBasic.Slides.Layout
Imports Aspose.Slides.Examples.VisualBasic.Slides.Media
Imports Aspose.Slides.Examples.VisualBasic.Slides.Notes
Imports Aspose.Slides.Examples.VisualBasic.Slides.Thumbnail
Imports Aspose.Slides.Examples.VisualBasic.Slides.Transitions
Imports Aspose.Slides.Examples.VisualBasic.SmartArts
Imports Aspose.Slides.Examples.VisualBasic.Tables
Imports Aspose.Slides.Examples.VisualBasic.Text
Imports Aspose.Slides.Examples.VisualBasic.VBA
Imports Microsoft.VisualBasic


Namespace Aspose.Slides.Examples.VisualBasic
    Class RunExamples
        <STAThread> _
        Public Shared Sub Main()
            Console.WriteLine("Open RunExamples.cs. " & vbLf & "In Main() method uncomment the example that you want to run.")
            Console.WriteLine("=====================================================")

            ' Uncomment the one you want to try out

            '''/ =====================================================
            '''/                    ActiveX
            '''/ =====================================================

            'ManageActiveXControl.Run()
            'LinkingVideoActiveXControl.Run()

            ' =====================================================
            '                      Charts
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
            'SetDataRange.Run()
            'SetInvertFillColorChart.Run()

            '''/ =====================================================
            '''/                    Presentations 
            '''/ =====================================================

            '''/ =====================================================
            '''/               Presentations - Conversion
            '''/ =====================================================

            'ConvertToPDF.Run()
            'ConvertToPDFWithHiddenSlides.Run()
            'CustomOptionsPDFConversion.Run()
            'ConvertPresentationToPasswordProtectedPDF.Run()
            'ConvertSpecificSlideToPDF.Run()
            'ConvertSlidesToPdfNotes.Run()
            'PresentationToTIFFWithDefaultSize.Run()
            'PresentationToTIFFWithCustomImagePixelFormat.Run()
            'ConvertWithNoteToTiff.Run()
            'ConvertWholePresentationToHTML.Run()
            'ConvertPresentationToResponsiveHTML.Run()
            'ConvertIndividualSlide.Run()
            'ConvertWithCustomSize.Run()
            'ConvertNotesSlideView.Run()
            'ConvertWithoutXpsOptions.Run()
            'ConvertWithXpsOptions.Run()
            'ConvetToSWF.Run()
            'ConversionToTIFFNotes.Run()
            'ConvertNotesSlideViewToPDF.Run()
            'CreateNewPresentation.Run()
            'PPTtoPPTX.Run()
            'ExportMediaFilestohtml.Run()

            ' =====================================================
            '''/ =====================================================
            '''/ Presentations -   Opening
            '''/ =====================================================
            '''/ =====================================================

            'OpenPresentation.Run()
            'OpenPasswordPresentation.Run()
            'VerifyingPresentationWithoutloading.Run()
            'GetFileFormat.Run()
            'GetRectangularCoordinatesofParagraph.Run()
            'GetPositionCoordinatesofPortion.Run()

            '''/ =====================================================
            '''/            Presentations -   Properties
            '''/ =====================================================

            'AccessBuiltinProperties.Run()
            'AccessModifyingProperties.Run()
            'AddCustomDocumentProperties.Run()
            'AccessOpenDoc.Run()
            'AccessProperties.Run()
            'ModifyBuiltinProperties.Run()
            'UpdatePresentationProperties.Run()
            'UpdatePresentationPropertiesUsingNewTemplate.Run()
            'UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate.Run()

            '''/ =====================================================
            '''/            Presentations -   Saving
            '''/ =====================================================

            'RemoveWriteProtection.Run()
            'SaveAsReadOnly.Run()
            'SaveProperties.Run()
            'SaveToFile.Run()
            'SaveToStream.Run()
            'SaveWithPassword.Run()
            'SaveAsPredefinedViewType.Run()

            '''/ =====================================================
            '''/                    Shapes
            '''/ =====================================================

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
            'FindShapeInSlide.Run()
            'FillShapeswithSolidColor.Run()
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
            'CreateScalingFactorThumbnail.Run()
            'CreateSmartArtChildNoteThumbnail.Run()
            'CreateGroupShape.Run()
            'AccessingAltTextinGroupshapes.Run()
            'CloneShapes.Run()
            'SetAlternativeText.Run()
            'RemoveShape.Run()
            'Hidingshapes.Run()
            'ChangeShapeOrder.Run()
            'ConnectShapesUsingConnectors.Run()
            'ConnectShapeUsingConnectionSite.Run()
            'ApplyBevelEffects.Run
            'AddVideoFrameFromWebSource.Run()


            '''/ =====================================================
            '''/                        Slides 
            '''/ =====================================================

            '''/ =====================================================
            '''/                    Slides - CRUD
            '''/ =====================================================

            'AccessSlides.Run()
            'AccessSlidebyIndex.Run()
            'AccessSlidebyID.Run()
            'CreateSlidesSVGImage.Run()
            'ChangePosition.Run()
            'CloneAtEndOfAnother.Run()
            'CloneAtEndOfAnotherSpecificPosition.Run()
            'CloneToAnotherPresentationWithMaster.Run()
            'CloneWithInSamePresentation.Run()
            'CloneWithinSamePresentationToEnd.Run()
            'CloneAnotherPresentationAtSpecifiedPosition.Run()
            'RemoveSlideUsingIndex.Run()
            'RemoveSlideUsingReference.Run()
            'AddSlides.Run()

            '''/ =====================================================
            '''/                    Slides - Notes
            '''/ =====================================================

            'RemoveNotesAtSpecificSlide.Run()
            'RemoveNotesFromAllSlides.Run()

            '''/ =====================================================
            '''/                    Slides - Background
            '''/ =====================================================

            'SetBackgroundToGradient.Run()
            'SetImageAsBackground.Run()
            'SetSlideBackgroundMaster.Run()
            'SetSlideBackgroundNormal.Run()

            '''/ =====================================================
            '''/                    Slides - Transitions
            '''/ =====================================================

            'BetterSlideTransitions.Run()
            'SimpleSlideTransitions.Run()
            'ManageSimpleSlideTransitions.Run()
            'ManagingBetterSlideTransitions.Run()
            'SetTransitionEffects.Run()

            '''/ =====================================================
            '''/                    Slides - Thumbnail
            '''/ =====================================================

            'ThumbnailFromSlide.Run()
            'ThumbnailFromSlideInNotes.Run()
            'ThumbnailWithUserDefinedDimensions.Run()

            '''/ =====================================================
            '''/                    Slides - Comments
            '''/ =====================================================

            'AddSlideComments.Run()
            'AccessSlideComments.Run()

            '''/ =====================================================
            '''/                    Slides - Layout
            '''/ =====================================================

            'AddLayoutSlides.Run()
            'SetSizeAndType.Run()
            'SetPDFPageSize.Run()

            '''/ =====================================================
            '''/                    Slides - HyperLink
            '''/ =====================================================

            'RemoveHyperlinks.Run()

            '''/ =====================================================
            '''/                    Slides - Media
            '''/ =====================================================

            'ExtractVideo.Run()

            '''/ =====================================================
            '''/            Rendering - Printing a Slide
            '''/ =====================================================

            'SetZoom.Run()
            'SetSlideNumber.Run()
            'DefaultPrinterPrinting.Run()
            'SpecificPrinterPrinting.Run()

            '''/ =====================================================
            '''/                    Smart Arts
            '''/ =====================================================

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
            'AccessSmartArtParticularLayout.Run()
            'ChangSmartArtShapeStyle.Run()
            'ChangeSmartArtShapeColorStyle.Run()
            'FillFormatSmartArtShapeNode.Run()
            'ChangeTextOnSmartArtNode.Run()
            'ChangeSmartArtLayout.Run()
            'CheckSmartArtHiddenProperty.Run()
            'ChangeSmartArtState.Run()
            'OrganizeChartLayoutType.Run()

            '''/ =====================================================
            '''/                    Tables
            '''/ =====================================================

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

            '''/ =====================================================
            '''/ Text
            '''/ =====================================================

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
            'ManageEmbeddedFonts.Run()

            '''/ =====================================================
            '''/                    VBA Macros
            '''/ =====================================================

            'AddVBAMacros.Run()
            'RemoveVBAMacros.Run()

            ' Stop before exiting

            Console.WriteLine(Constants.vbLf + Constants.vbLf & "Program Finished. Press any key to exit....")
            Console.ReadKey()

        End Sub

        Protected Sub Page_Load(sender As Object, e As EventArgs)


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

        Public Shared Function GetDataDir_Conversion() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Presentations/Conversion/"))
        End Function

        Public Shared Function GetDataDir_PresentationProperties() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Presentations/Properties/"))
        End Function

        Public Shared Function GetDataDir_PresentationSaving() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Presentations/Saving/"))
        End Function

        Public Shared Function GetDataDir_PresentationOpening() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Presentations/Opening/"))
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

        Public Shared Function GetDataDir_Slides_Presentations_CRUD() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/CRUD/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Notes() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Notes/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Background() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Background/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Transitions() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Transitions/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Thumbnail() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Thumbnail/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Comments() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Comments/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Layout() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Layout/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Hyperlink() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Hyperlink/"))
        End Function

        Public Shared Function GetDataDir_Slides_Presentations_Media() As [String]
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Slides/Media/"))
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