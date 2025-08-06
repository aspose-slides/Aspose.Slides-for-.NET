using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells.Drawing;

/*
This example demonstrates how to programmatically access, set, remove, and list font mappings for different script 
tags (e.g., Latn, Cyrl, Jpan) following the BCP-47 script code standard. 
*/

namespace CSharp.Text
{
    class ManageScriptFontsExample
    {
        public static void Run()
        {
            using (Presentation pres = new Presentation())
            {
                // Get all script mapping
                IDictionary<string, string> scriptFontMap = pres.MasterTheme.FontScheme.Major.GetScriptFontMap();
                foreach (var kvp in scriptFontMap)
                {
                    Console.WriteLine(kvp.Key + " = " + kvp.Value);
                }

                // Get script font
                Console.WriteLine("Font for \"Thaa\" tag is " + pres.MasterTheme.FontScheme.Major.GetScriptFont("Thaa"));

                // Set script font
                pres.MasterTheme.FontScheme.Major.SetScriptFont("Thaa", "Super Thaa");
                pres.MasterTheme.FontScheme.Minor.RemoveScriptFont("Geor");

                // Check script font
                Console.WriteLine("Font for \"Thaa\" tag is " + pres.MasterTheme.FontScheme.Major.GetScriptFont("Thaa"));
            }
        }
    }
}
