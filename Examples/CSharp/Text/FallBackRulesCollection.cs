using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class FallBackRulesCollection
    {
        public static void Run()
        {

            //ExStart:FallBackRulesCollection

            using (Presentation presentation = new Presentation())
            {
                IFontFallBackRulesCollection userRulesList = new FontFallBackRulesCollection();

                userRulesList.Add(new FontFallBackRule(0x0B80, 0x0BFF, "Vijaya"));
                userRulesList.Add(new FontFallBackRule(0x3040, 0x309F, "MS Mincho, MS Gothic"));

                presentation.FontsManager.FontFallBackRulesCollection = userRulesList;
            }
            //ExEnd:FallBackRulesCollection

        }
    }
}
