using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class SetFontFallBack
    {
        public static void Run() {

            //ExStart:SetFontFallBack

            uint startUnicodeIndex = 0x0B80;
            uint endUnicodeIndex = 0x0BFF;

            IFontFallBackRule firstRule = new FontFallBackRule(startUnicodeIndex, endUnicodeIndex, "Vijaya");
            IFontFallBackRule secondRule = new FontFallBackRule(0x3040, 0x309F, "MS Mincho, MS Gothic");

            //Also the fonts list can be added in several ways:
            string[] fontNames = new string[] { "Segoe UI Emoji, Segoe UI Symbol", "Arial" };

            IFontFallBackRule thirdRule = new FontFallBackRule(0x1F300, 0x1F64F, fontNames);

            //ExEnd:SetFontFallBack

        }
    }
}
