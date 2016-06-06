//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace CSharp.Charts
{
    public class NumberFormat
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate the presentation//Instantiate the presentation
            Presentation pres = new Presentation();

            //Access the first presentation slide
            ISlide slide = pres.Slides[0];

            //Adding a defautlt clustered column chart
            IChart chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400);

            //Accessing the chart series collection
            IChartSeriesCollection series = chart.ChartData.Series;

            //Setting the preset number format
            //Traverse through every chart series
            foreach (ChartSeries ser in series)
            {
                //Traverse through every data cell in series
                foreach (IChartDataPoint cell in ser.DataPoints)
                {
                    //Setting the number format
                    cell.Value.AsCell.PresetNumberFormat = 10; //0.00%
                }
            }

            // Saving presentation
            pres.Save(dataDir + "PresetNumberFormat.pptx", SaveFormat.Pptx);
        }
    }
}