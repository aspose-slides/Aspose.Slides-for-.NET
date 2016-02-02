// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using Aspose.Slides.Util;
namespace Find_and_Replace
{
    class Program
    {
        static void Main(string[] args)
        {
            findReplaceText("Some Text", "Replaced Text");
        }
        public static void findReplaceText(string strToFind, string strToReplaceWith)
        {
            string MyDir = @"Files\";
            //Open the presentation
            Presentation pres = new Presentation(MyDir + "Find and Replace.ppt");
            //Get all text boxes in the presentation
            ITextBox[] tb = SlideUtil.GetAllTextBoxes(pres, false);
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
            pres.Write(MyDir + "Find and Replace.pptx");

        }
    }
}
