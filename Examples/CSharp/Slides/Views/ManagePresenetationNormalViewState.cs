using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Views
{
    class ManagePresenetationNormalViewState
    {
        public static void Run() {

            //ExStart:ManagePresenetationNormalViewState
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Views();

            using (Presentation pres = new Presentation())
            {
                pres.ViewProperties.NormalViewProperties.HorizontalBarState = SplitterBarStateType.Restored;
                pres.ViewProperties.NormalViewProperties.VerticalBarState = SplitterBarStateType.Maximized;

                pres.ViewProperties.NormalViewProperties.RestoredTop.AutoAdjust = true;
                pres.ViewProperties.NormalViewProperties.RestoredTop.DimensionSize = 80;
                pres.ViewProperties.NormalViewProperties.ShowOutlineIcons = true;

                pres.Save(dataDir+ "presentation_normal_view_state.pptx", SaveFormat.Pptx);
            }

            //ExEnd:ManagePresenetationNormalViewState
        }
    }
}
