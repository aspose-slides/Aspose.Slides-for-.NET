using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using CSharp.Charts;
using CSharp.Presentations;
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

            ChartEntities.Run();
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
            //AccessOpenDoc.Run();
            //AccessProperties.Run();
            //ConvertToPDF.Run();
            //ConvertWithNoteToTiff.Run();
            //Convert_HTML.Run();
            //Convert_Tiff_Custom.Run();
            //Convert_Tiff_Default.Run();
            //Convert_XPS.Run();
            //Convert_XPS_Options.Run();
            //ModifyBuiltinProperties.Run();
            //OpenPasswordPresentation.Run();
            //OpenPresentation.Run();
            //PPTtoPPTX.Run();
            //RemoveWriteProtection.Run();
            //SaveAsReadOnly.Run();
            //SaveProperties.Run();
            //SaveToFile.Run();
            //SaveWithPassword.Run();

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
            return Path.GetFullPath("../../ProgrammersGuide/Slides-Presentations/Data/");
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