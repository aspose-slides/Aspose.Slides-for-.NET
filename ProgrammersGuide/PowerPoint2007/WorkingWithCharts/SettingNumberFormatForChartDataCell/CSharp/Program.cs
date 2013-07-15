//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using Aspose.Slides.Pptx.Charts;

namespace SettingNumberFormatForChartDataCell
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate the presentation
            PresentationEx pres = new PresentationEx();

            //Access the first presentation slide
            SlideEx slide = pres.Slides[0];

            //Adding a defautlt clustered column chart
            ChartEx chart = slide.Shapes.AddChart(ChartTypeEx.ClusteredColumn, 50, 50, 500, 400);

            //Accessing the chart series collection
            ChartSeriesExCollection series = chart.ChartData.Series;


            //Setting the preset number format

            //Traverse through every chart series
            foreach (ChartSeriesEx ser in series)
            {
                //Traverse through every data cell in series
                foreach (ChartDataCell cell in ser.Values)
                {
                    //Setting the number format  
                    cell.PresetNumberFormat = 10; //0.00%
                }
            }

            //Saving presentation
            pres.Write(dataDir + "PresetNumberFormat.pptx");


            //Now setting the custom number format

            //Traverse through every chart series
            foreach (ChartSeriesEx ser in series)
            {
                //Traverse through every data cell in series
                foreach (ChartDataCell cell in ser.Values)
                {
                    //Setting the number format  
                    cell.CustomNumberFormat = "0.00000";
                }
            }
            //Saving presentation
            pres.Write(dataDir + "CustomNumberFormat.pptx");

        }
    }
}