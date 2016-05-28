/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

using System;
using System.IO;
using CSharp.Charts;
using CSharp.Presentations;
using CSharp.ProgrammersGuide.Presentations;
using CSharp.Shapes;
using CSharp.Slides;
using CSharp.SmartArts;
using CSharp.Tables;
using CSharp.Text;

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
            //ConvetToSWF.Run();
            //ConversionToTIFFNotes.Run();
            //ConvertNotesSlideViewToPDF.Run();
            //CreateNewPresentation.Run();
           
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
            //FormatJoinStyles.Run();
            //FormatLines.Run();
            //FormattedEllipse.Run();
            //FormattedRectangle.Run();
            //PictureFrameFormatting.Run();
            //RotatingShapes.Run();
            //SimpleEllipse.Run();
            //SimpleRectangle.Run();


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

            //// =====================================================
            //// =====================================================
            //// Tables
            //// =====================================================
            //// =====================================================

            //RemovingRowColumn.Run();
            //TableFromScratch.Run();
            //TableWithCellBorders.Run();
            //UpdateExistingTable.Run();

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


            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        protected void Page_Load(object sender, EventArgs e)
        {


        }

        public static String GetDataDir_Charts()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Charts/Data/");
        }

        public static String GetDataDir_Presentations()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Presentations/Data/");
        }

        public static String GetDataDir_Shapes()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Shapes/Data/");
        }

        public static String GetDataDir_Slides_Presentations()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Slides-Presentation/Data/");
        }

        public static String GetDataDir_SmartArts()
        {
            return Path.GetFullPath("../../ProgrammersGuide/SmartArts/Data/");
        }

        public static String GetDataDir_Tables()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Tables/Data/");
        }

        public static String GetDataDir_Text()
        {
            return Path.GetFullPath("../../ProgrammersGuide/Text/Data/");
        }
    }
}