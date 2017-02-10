using System;
using System.IO;
using Aspose.Slides.Examples.CSharp.ActiveX;
using Aspose.Slides.Examples.CSharp.Charts;
using Aspose.Slides.Examples.CSharp.Conversion;
using Aspose.Slides.Examples.CSharp.Presentations;
using Aspose.Slides.Examples.CSharp.Presentations.Conversion;
using Aspose.Slides.Examples.CSharp.Presentations.Opening;
using Aspose.Slides.Examples.CSharp.Presentations.Properties;
using Aspose.Slides.Examples.CSharp.Presentations.Saving;
using Aspose.Slides.Examples.CSharp.Rendering.Printing;
using Aspose.Slides.Examples.CSharp.Shapes;
using Aspose.Slides.Examples.CSharp.Slides;
using Aspose.Slides.Examples.CSharp.Slides.Background;
using Aspose.Slides.Examples.CSharp.Slides.Comments;
using Aspose.Slides.Examples.CSharp.Slides.CRUD;
using Aspose.Slides.Examples.CSharp.Slides.Hyperlink;
using Aspose.Slides.Examples.CSharp.Slides.Layout;
using Aspose.Slides.Examples.CSharp.Slides.Media;
using Aspose.Slides.Examples.CSharp.Slides.Notes;
using Aspose.Slides.Examples.CSharp.Slides.Thumbnail;
using Aspose.Slides.Examples.CSharp.Slides.Transitions;
using Aspose.Slides.Examples.CSharp.SmartArts;
using Aspose.Slides.Examples.CSharp.Tables;
using Aspose.Slides.Examples.CSharp.Text;
using Aspose.Slides.Examples.CSharp.VBA;


namespace Aspose.Slides.Examples.CSharp
{
    class RunExamples
    {
        [STAThread]
        public static void Main()
        {
            Console.WriteLine("Open RunExamples.cs. \nIn Main() method uncomment the example that you want to run.");
            Console.WriteLine("=====================================================");

            // Uncomment the one you want to try out

            //// =====================================================
            ////                    ActiveX
            //// =====================================================

            //ManageActiveXControl.Run();
            //LinkingVideoActiveXControl.Run();

            // =====================================================
            //                      Charts
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
            SetCategoryAxisLabelDistance.Run();
            //SetlegendCustomOptions.Run();
            //SetDataLabelsPercentageSign.Run();
            //DoughnutChartHole.Run();
            //ManagePropertiesCharts.Run();
            //SetGapWidth.Run();
            //AutomaticChartSeriescolor.Run();
            //DisplayPercentageAsLabels.Run();
            //SecondPlotOptionsforCharts.Run();
            //SetMarkerOptions.Run();
            //SetDataRange.Run();
         //UsingWorkBookChartcellAsDatalabel.Run();
            //// =====================================================
            ////                    Presentations 
            //// =====================================================

            //// =====================================================
            ////               Presentations - Conversion
            //// =====================================================

            //ConvertToPDF.Run();
            //ConvertToPDFWithHiddenSlides.Run();
            //CustomOptionsPDFConversion.Run();
            //ConvertPresentationToPasswordProtectedPDF.Run();
            //ConvertSpecificSlideToPDF.Run();
            //ConvertSlidesToPdfNotes.Run();
            //PresentationToTIFFWithDefaultSize.Run();
            //PresentationToTIFFWithCustomImagePixelFormat.Run();
            //ConvertWithNoteToTiff.Run();
            //ConvertWholePresentationToHTML.Run();
            //ConvertPresentationToResponsiveHTML.Run();
            //ConvertIndividualSlide.Run();
            //ConvertWithCustomSize.Run();
            //ConvertNotesSlideView.Run();
            //ConvertWithoutXpsOptions.Run();
            //ConvertWithXpsOptions.Run();
            //ConvetToSWF.Run();
            //ConversionToTIFFNotes.Run();
            //ConvertNotesSlideViewToPDF.Run();
            //CreateNewPresentation.Run();
            //PPTtoPPTX.Run();
            //ExportMediaFilestohtml.Run();
            //SetInvertFillColorChart.Run();

            // =====================================================
            //// =====================================================
            //// Presentations -   Opening
            //// =====================================================
            //// =====================================================

            //OpenPresentation.Run();
            //OpenPasswordPresentation.Run();
            //VerifyingPresentationWithoutloading.Run();
            //GetFileFormat.Run();
            //GetRectangularCoordinatesofParagraph.Run();
            //GetPositionCoordinatesofPortion.Run();

            //// =====================================================
            ////            Presentations -   Properties
            //// =====================================================

            //AccessBuiltinProperties.Run();
            //AccessModifyingProperties.Run();
            //AddCustomDocumentProperties.Run();
            //AccessOpenDoc.Run();
            //AccessProperties.Run();
            //ModifyBuiltinProperties.Run();
            //UpdatePresentationProperties.Run();
            //UpdatePresentationPropertiesUsingNewTemplate.Run();
            //UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate.Run();

            //// =====================================================
            ////            Presentations -   Saving
            //// =====================================================

            //RemoveWriteProtection.Run();
            //SaveAsReadOnly.Run();
            //SaveProperties.Run();
            //SaveToFile.Run();
            //SaveToStream.Run();
            //SaveWithPassword.Run();
            //SaveAsPredefinedViewType.Run();

            //// =====================================================
            ////                    Shapes
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
            //AddVideoFrameFromWebSource.Run();

            //// =====================================================
            ////                        Slides 
            //// =====================================================

            //// =====================================================
            ////                    Slides - CRUD
            //// =====================================================

            //AccessSlides.Run();
            //AccessSlidebyIndex.Run();
            //AccessSlidebyID.Run();
            //CreateSlidesSVGImage.Run();
            //ChangePosition.Run();
            //CloneAtEndOfAnother.Run();
            //CloneAtEndOfAnotherSpecificPosition.Run();
            //CloneToAnotherPresentationWithMaster.Run();
            //CloneWithInSamePresentation.Run();
            //CloneWithinSamePresentationToEnd.Run();
            //CloneAnotherPresentationAtSpecifiedPosition.Run();
            //RemoveSlideUsingIndex.Run();
            //RemoveSlideUsingReference.Run();
            //AddSlides.Run();

            //// =====================================================
            ////                    Slides - Notes
            //// =====================================================

            //RemoveNotesAtSpecificSlide.Run();
            //RemoveNotesFromAllSlides.Run();           

            //// =====================================================
            ////                    Slides - Background
            //// =====================================================

            //SetBackgroundToGradient.Run();
            //SetImageAsBackground.Run();
            //SetSlideBackgroundMaster.Run();
            //SetSlideBackgroundNormal.Run();

            //// =====================================================
            ////                    Slides - Transitions
            //// =====================================================

            //BetterSlideTransitions.Run();
            //SimpleSlideTransitions.Run();
            //ManageSimpleSlideTransitions.Run();
            //ManagingBetterSlideTransitions.Run();
            //SetTransitionEffects.Run();

            //// =====================================================
            ////                    Slides - Thumbnail
            //// =====================================================

            //ThumbnailFromSlide.Run();
            //ThumbnailFromSlideInNotes.Run();
            //ThumbnailWithUserDefinedDimensions.Run();

            //// =====================================================
            ////                    Slides - Comments
            //// =====================================================

            //AddSlideComments.Run();
            //AccessSlideComments.Run();

            //// =====================================================
            ////                    Slides - Layout
            //// =====================================================

            //AddLayoutSlides.Run();
            //SetSizeAndType.Run();
            //SetPDFPageSize.Run();       
            //SetSlideSizeScale.Run();
            //// =====================================================
            ////                    Slides - HyperLink
            //// =====================================================

            //RemoveHyperlinks.Run();

            //// =====================================================
            ////                    Slides - Media
            //// =====================================================

            //ExtractVideo.Run();

            //// =====================================================
            ////            Rendering - Printing a Slide
            //// =====================================================

            //SetZoom.Run();
            //SetSlideNumber.Run();
            //DefaultPrinterPrinting.Run();
            //SpecificPrinterPrinting.Run();

            //// =====================================================
            ////                    Smart Arts
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
            ////                    Tables
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
            //TextFormattingInsideTableColumn.Run();
            //TextFormattingInsideTableRow.Run();
            //SetFormattingInsideTable.Run();
            //// =====================================================
            //// Text
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
            //ManageParagraphPictureBulletsInPPT.Run();
            //ManageEmbeddedFonts.Run();

            //// =====================================================
            ////                    VBA Macros
            //// =====================================================

            //AddVBAMacros.Run();
            //RemoveVBAMacros.Run();

            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();


        }

        protected void Page_Load(object sender, EventArgs e)
        {}

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

        public static String GetDataDir_Conversion()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Presentations/Conversion/");
        }

        public static String GetDataDir_PresentationProperties()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Presentations/Properties/");
        }

        public static String GetDataDir_PresentationSaving()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Presentations/Saving/");
        }

        public static String GetDataDir_PresentationOpening()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Presentations/Opening/");
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

        public static String GetDataDir_Slides_Presentations_CRUD()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/CRUD/");
        }

        public static String GetDataDir_Slides_Presentations_Notes()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Notes/");
        }

        public static String GetDataDir_Slides_Presentations_Background()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Background/");
        }

        public static String GetDataDir_Slides_Presentations_Transitions()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Transitions/");
        }

        public static String GetDataDir_Slides_Presentations_Thumbnail()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Thumbnail/");
        }

        public static String GetDataDir_Slides_Presentations_Comments()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Comments/");
        }

        public static String GetDataDir_Slides_Presentations_Layout()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Layout/");
        }

        public static String GetDataDir_Slides_Presentations_Hyperlink()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Hyperlink/");
        }

        public static String GetDataDir_Slides_Presentations_Media()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Slides/Media/");
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