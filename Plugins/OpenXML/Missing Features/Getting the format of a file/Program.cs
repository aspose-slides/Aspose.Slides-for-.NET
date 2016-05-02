using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Getting_the_format_of_a_file
{
    class Program
    {
        static void Main(string[] args)
        {
            string Path = @"E:\Aspose\Aspose Vs OpenXML\Files\";
            IPresentationInfo info;
            info = PresentationFactory.Instance.GetPresentationInfo(Path + "Test.pptx");


            switch (info.LoadFormat)
            {
                case LoadFormat.Pptx:
                    {
                        break;
                    }
                case LoadFormat.Unknown:
                    {
                        break;
                    }
            }
        }
    }
}
