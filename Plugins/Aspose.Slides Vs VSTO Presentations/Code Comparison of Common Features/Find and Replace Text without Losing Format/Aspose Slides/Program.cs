using Aspose.Slides;
using Aspose.Slides.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            findReplaceText("d", "Aspose");
        }
        private static void findReplaceText(string strToFind, string strToReplaceWith)
        {

            //Open the presentation
            Presentation pres = new Presentation("mytextone.ppt");
            //Get all text boxes in the presentation
            ITextBox[] tb = PresentationScanner.GetAllTextBoxes(pres, false);
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
            pres.Write("myTextOneAspose.ppt");
        }
    }
}
