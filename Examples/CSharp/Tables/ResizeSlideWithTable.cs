using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Tables
{
    class ResizeSlideWithTable
    {
        public static void Run() {

            //ExStart:ResizeSlideWithTable
            Presentation presentation = new Presentation("D:\\Test.pptx");

            //Old slide size
            float currentHeight = presentation.SlideSize.Size.Height;
            float currentWidth = presentation.SlideSize.Size.Width;

            //Changing slide size
            presentation.SlideSize.Type = SlideSizeType.A4Paper;
            //presentation.SlideSize.Orientation = SlideOrienation.Portrait;

            //New slide size
            float newHeight = presentation.SlideSize.Size.Height;
            float newWidth = presentation.SlideSize.Size.Width;


            float ratioHeight = newHeight / currentHeight;
            float ratioWidth = newWidth / currentWidth;

            foreach (IMasterSlide master in presentation.Masters)
            {
                foreach (IShape shape in master.Shapes)
                {
                    //Resize position
                    shape.Height = shape.Height * ratioHeight;
                    shape.Width = shape.Width * ratioWidth;

                    //Resize shape size if required 
                    shape.Y = shape.Y * ratioHeight;
                    shape.X = shape.X * ratioWidth;

                }

                foreach (ILayoutSlide layoutslide in master.LayoutSlides)
                {
                    foreach (IShape shape in layoutslide.Shapes)
                    {
                        //Resize position
                        shape.Height = shape.Height * ratioHeight;
                        shape.Width = shape.Width * ratioWidth;

                        //Resize shape size if required 
                        shape.Y = shape.Y * ratioHeight;
                        shape.X = shape.X * ratioWidth;

                    }

                }
            }

            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    //Resize position
                    shape.Height = shape.Height * ratioHeight;
                    shape.Width = shape.Width * ratioWidth;

                    //Resize shape size if required 
                    shape.Y = shape.Y * ratioHeight;
                    shape.X = shape.X * ratioWidth;
                    if (shape is ITable)
                    {
                        ITable table = (ITable)shape;
                        foreach (IRow row in table.Rows)
                        {
                            row.MinimalHeight = row.MinimalHeight * ratioHeight;
                            //   row.Height = row.Height * ratioHeight;
                        }
                        foreach (IColumn col in table.Columns)
                        {
                            col.Width = col.Width * ratioWidth;

                        }
                    }

                }
            }

            presentation.Save("D:\\Resize.pptx", SaveFormat.Pptx);
            //ExEnd:ResizeSlideWithTable

        }
    }
}
