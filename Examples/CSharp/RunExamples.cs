
using System;
using System.IO;
using CSharp.Charts;
using CSharp.Presentations;
using CSharp.Presentations;
using CSharp.Rendering.Printing;
using CSharp.Shapes;
using CSharp.SmartArts;
using CSharp.Shapes;
using CSharp.Slides;
using CSharp.SmartArts;
using CSharp.Tables;
using CSharp.Text;
using CSharp.VBA;

namespace CSharp
{
    class RunExamples
    {
        [STAThread]
        public static void Main()
        {
            Console.WriteLine("Open RunExamples.cs. In Main() method, Un-comment the example that you want to run");
            Console.WriteLine("=====================================================");

            // Un-comment the one you want to try out

            //// =====================================================
            //// =====================================================
            //// ActiveX
            //// =====================================================
            //// =====================================================

            ManageActiveXControl.Run();
            //LinkingVideoActiveXControl.Run();

            // =====================================================
            // =====================================================
            // Charts
            // =====================================================
            // =====================================================

            //ChartEntities.Run();
            //ChartTrendLines.Run();
            //ExistingChart.Run();
            //NormalCharts.Run();
            //NumberFormat.Run();
            //ScatteredChart.Run();
            //PieChart.Run();
            //ChangeChartCategoryAxis.Run();
            //DisplayChartLabels.Run();
            //AddErrorBars.Run();
            //AddCustomError.Run();
            //AnimatingSeries.Run();
            //AnimatingSeriesElements.Run();
            //AnimatingCategoriesElements.Run();
            //SetChartSeriesOverlap.Run();
            //SetAutomaticSeriesFillColor.Run();
            //SetCategoryAxisLabelDistance.Run();
            //SetlegendCustomOptions.Run();
            //SetDataLabelsPercentageSign.Run();
            //DoughnutChartHole.Run();
            //ManagePropertiesCharts.Run();
            //SetGapWidth.Run();
            //AutomaticChartSeriescolor.Run();
            //DisplayPercentageAsLabels.Run();
            //SecondPlotOptionsforCharts.Run();
            //SetMarkerOptions.Run();

            //// =====================================================
            //// =====================================================
            //// Presentations
            //// =====================================================
            //// =====================================================

            //AccessBuiltinProperties.Run();
            //AccessModifyingProperties.Run();
            //AddCustomDocumentProperties.Run();
            //AccessOpenDoc.Run();
            //AccessProperties.Run();
            //ConvertToPDF.Run();
            //CustomOptionsPDFConversion.Run();
            //ConvertPresentationToPasswordProtectedPDF.Run();
            //ConvertSpecificSlideToPDF.Run();
            //ConvertSlidesToPdfNotes.Run();
            //PresentationToTIFFWithDefaultSize.Run();
            //PresentationToTIFFWithCustomImagePixelFormat.Run();
            //ConvertWithNoteToTiff.Run();
            //Convert_HTML.Run();
            //ConvertIndividualSlide.Run();
            //Convert_Tiff_Custom.Run();
            //Convert_Tiff_Default.Run();
            //Convert_XPS.Run();
            //Convert_XPS_Options.Run();
            //ModifyBuiltinProperties.Run();
            //OpenPresentation.Run();
            //OpenPasswordPresentation.Run();
            //PPTtoPPTX.Run();
            //RemoveWriteProtection.Run();
            //SaveAsReadOnly.Run();
            //SaveProperties.Run();
            //SaveToFile.Run();
            //SaveToStream.Run();
            //SaveWithPassword.Run();
            //SaveAsPredefinedViewType.Run();
            //VerifyingPresentationWithoutloading.Run();
            //ExportMediaFilestohtml.Run();
            //GetFileFormat.Run();
            //ConversionToTIFFNotes.Run();
            //ConvertNotesSlideViewToPDF.Run();
            //CreateNewPresentation.Run();
            //ConvetToSWF.Run();
            //GetRectangularCoordinatesofParagraph.Run();
            //GetPositionCoordinatesofPortion.Run();

            //// =====================================================
            //// =====================================================
            //// Shapes
            //// =====================================================
            //// =====================================================

            //AccessOLEObjectFrame.Run();
            //AddArrowShapedLine.Run();
            //AddArrowShapedLineToSlide.Run();
            //AddAudioFrame.Run();
            //AddOLEObjectFrame.Run();
            //AddPlainLineToSlide.Run();
            //AddSimplePictureFrames.Run();
            //AddVideoFrame.Run();
            //AnimationsOnShapes.Run();
            //ChangeOLEObjectData.Run();
            //ConnectorLineAngle.Run();
            //EmbeddedVideoFrame.Run();
            //FillShapesGradient.Run();
            //FillShapesPattern.Run();
            //FillShapesPicture.Run();
            //FindShapeInSlide.Run();
            //FillShapeswithSolidColor.Run();
            //FormatJoinStyles.Run();
            //FormatLines.Run();
            //FormattedEllipse.Run();
            //FormattedRectangle.Run();
            //PictureFrameFormatting.Run();
            //RotatingShapes.Run();
            //SimpleEllipse.Run();
            //SimpleRectangle.Run();
            //AddRelativeScaleHeightPictureFrame.Run();
            //CreateShapeThumbnail.Run();
            //CreateBoundsShapeThumbnail.Run();
            //CreateScalingFactorThumbnail.Run();
            //CreateSmartArtChildNoteThumbnail.Run();
            //CreateGroupShape.Run();
            //AccessingAltTextinGroupshapes.Run();
            //CloneShapes.Run();
            //SetAlternativeText.Run();
            //RemoveShape.Run();
            //Hidingshapes.Run();
            //ChangeShapeOrder.Run();
            //ConnectShapesUsingConnectors.Run();
            //ConnectShapeUsingConnectionSite.Run();
            //ApplyBevelEffects.Run();

            //// =====================================================
            //// =====================================================
            //// Slides in Presentation
            //// =====================================================
            //// =====================================================

            //AccessSlides.Run();
            //AddSlides.Run();
            //BetterSlideTransitions.Run();
            //ChangePosition.Run();
            //CloneAtEndOfAnother.Run();
            //CloneAtEndOfAnotherSpecificPosition.Run();
            //CloneToAnotherPresentationWithMaster.Run();
            //CloneWithInSamePresentation.Run();
            //CloneWithinSamePresentationToEnd.Run();
            //CreateSlidesSVGImage.Run();
            //RemoveSlideUsingIndex.Run();
            //RemoveSlideUsingReference.Run();
            //SetBackgroundToGradient.Run();
            //SetImageAsBackground.Run();
            //SetSlideBackgroundMaster.Run();
            //SetSlideBackgroundNormal.Run();
            //SimpleSlideTransitions.Run();
            //ThumbnailFromSlide.Run();
            //ThumbnailFromSlideInNotes.Run();
            //ThumbnailWithUserDefinedDimensions.Run();
            //AccessSlidebyIndex.Run();
            //AccessSlidebyID.Run();
            //CloneAnotherPresentationAtSpecifiedPosition.Run();
            //ManagSimpleSlideTransitions.Run();
            //ManagingBetterSlideTransitions.Run();
            //AddSlideComments.Run();
            //AccessSlideComments.Run();
            //RemoveHyperlinks.Run();
            //AddLayoutSlides.Run();
            //SetSizeAndType.Run();
            //SetPDFPageSize.Run();
            //RemoveNotesAtSpecificSlide.Run();
            //RemoveNotesFromAllSlides.Run();
            //ExtractVideo.Run();
            //SetTransitionEffects.Run();

            //// =====================================================
            //// =====================================================
            //// Rendering - Printing a Slide
            //// =====================================================
            //// =====================================================

            //SetZoom.Run();
            //SetSlideNumber.Run();
            //DefaultPrinterPrinting.Run();
            //SpecificPrinterPrinting.Run();

            //// =====================================================
            //// =====================================================
            //// Smart Arts
            //// =====================================================
            //// =====================================================

            //AccessChildNodes.Run();
            //AccessChildNodeSpecificPosition.Run();
            //AccessSmartArt.Run();
            //AccessSmartArtShape.Run();
            //AddNodes.Run();
            //AddNodesSpecificPosition.Run();
            //AssistantNode.Run();
            //CreateSmartArtShape.Run();
            //RemoveNode.Run();
            //RemoveNodeSpecificPosition.Run();
            //SmartArtNodeLevel.Run();
            //AccessSmartArtParticularLayout.Run();
            //ChangSmartArtShapeStyle.Run();
            //ChangeSmartArtShapeColorStyle.Run();
            //FillFormatSmartArtShapeNode.Run();
            //ChangeTextOnSmartArtNode.Run();
            //ChangeSmartArtLayout.Run();
            //CheckSmartArtHiddenProperty.Run();
            //ChangeSmartArtState.Run();
            //OrganizeChartLayoutType.Run();

            //// =====================================================
            //// =====================================================
            //// Tables
            //// =====================================================
            //// =====================================================

            //RemovingRowColumn.Run();
            //TableFromScratch.Run();
            //TableWithCellBorders.Run();
            //UpdateExistingTable.Run();
            //AddImageinsideTableCell.Run();
            //CloningInTable.Run();
            //VerticallyAlignText.Run();
            //StandardTables.Run();
            //MergeCells.Run();
            //MergeCell.Run();
            //CellSplit.Run();

            //// =====================================================
            //// =====================================================
            //// Text
            //// =====================================================
            //// =====================================================

            //DefaultFonts.Run();
            //ExportingHTMLText.Run();
            //FontFamily.Run();
            //FontProperties.Run();
            //ImportingHTMLText.Run();
            //MultipleParagraphs.Run();
            //ParagraphBullets.Run();
            //ParagraphIndent.Run();
            //ParagraphsAlignment.Run();
            //ReplacingText.Run();
            //ShadowEffects.Run();
            //TextBoxHyperlink.Run();
            //TextBoxOnSlideProgram.Run();
            //ApplyInnerShadow.Run();
            //ManageParagraphFontProperties.Run();
            //SetTextFontProperties.Run();
            //ReplaceFontsExplicitly.Run();
            //RuleBasedFontsReplacement.Run();
            //SetAutofitOftextframe.Run();
            //SetAnchorOfTextFrame.Run();
            //RotatingText.Run();
            //LineSpacing.Run();
            //ApplyOuterShadow.Run();
            //CustomRotationAngleTextframe.Run();
            //UseCustomFonts.Run(); 

            //// =====================================================
            //// =====================================================
            //// VBA Macros
            //// =====================================================
            //// =====================================================

            //AddVBAMacros.Run();
            //RemoveVBAMacros.Run();

            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();


        }

        protected void Page_Load(object sender, EventArgs e)
        {


        }

        public static String GetDataDir_ActiveX()
        {
            return Path.GetFullPath(GetDataDir_Data() + "ActiveX/");
        }
        public static String GetDataDir_Charts()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Charts/");
        }
        public static String GetDataDir_VBA()
        {
            return Path.GetFullPath(GetDataDir_Data() + "VBA/");
        }
        public static String GetDataDir_Presentations()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Presentations/");
        }

        public static String GetDataDir_Rendering()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Rendering-Printing/");
        }

        public static String GetDataDir_Shapes()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Shapes/");
        }

        public static String GetDataDir_Slides_Presentations()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/");
        }

        public static String GetDataDir_SmartArts()
        {
            return Path.GetFullPath(GetDataDir_Data() + "SmartArts/");
        }

        public static String GetDataDir_Tables()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Tables/");
        }

        public static String GetDataDir_Text()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Text/");
        }

        public static String GetDataDir_Video()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Video/");
        }

        private static string GetDataDir_Data()
        {
            var parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            string startDirectory = null;
            if (parent != null)
            {
                var directoryInfo = parent.Parent;
                if (directoryInfo != null)
                {
                    startDirectory = directoryInfo.FullName;
                }
            }
            else
            {
                startDirectory = parent.FullName;
            }
            return startDirectory != null ? Path.Combine(startDirectory, "Data\\") : null;
        }



    }
}