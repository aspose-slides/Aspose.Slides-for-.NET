// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using Aspose.Slides.Util;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            findReplaceText("This", "Replaced Text");
        }
        public static void findReplaceText(string strToFind, string strToReplaceWith)
        {
            string FilePath = @"..\..\..\Sample Files\";
            //Open the presentation
            Presentation pres = new Presentation(FilePath + "Find and Replace.pptx");
            //Get all text boxes in the presentation
            ITextFrame[] tb = SlideUtil.GetAllTextBoxes(pres.Slides[0]);
            for (int i = 0; i < tb.Length; i++)
                foreach (Paragraph para in tb[i].Paragraphs)
                    foreach (Portion port in para.Portions)
                        //Find text to be replaced
                        if (port.Text.Contains(strToFind))
                        //Replace exisitng text with the new text
                        {
                            string str = port.Text;
                            int idx = str.IndexOf(strToFind);
                            string strStartText = str.Substring(0, idx);
                            string strEndText = str.Substring(idx + strToFind.Length, str.Length - 1 - (idx + strToFind.Length - 1));
                            port.Text = strStartText + strToReplaceWith + strEndText;
                        }
            pres.Save(FilePath + "Find and Replace.pptx",Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}
