﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.CAD.Examples.CSharp.ApplyLicense
{
    class MeteredLicensing
    {

        public static void Run()
        {
            //ExStart:MeteredLicensing

            // Create an instance of CAD Metered class
            Aspose.Slides.Metered metered = new Aspose.Slides.Metered();

            // Access the setMeteredKey property and pass public and private keys as parameters
            metered.SetMeteredKey("*****", "*****");

            // Get metered data amount before calling API
            decimal amountbefore = Aspose.Slides.Metered.GetConsumptionQuantity();

            // Display information
            Console.WriteLine("Amount Consumed Before: " + amountbefore.ToString());
            // Get metered data amount After calling API
            decimal amountafter = Aspose.Slides.Metered.GetConsumptionQuantity();

            // Display information
            Console.WriteLine("Amount Consumed After: " + amountafter.ToString());

            // Print status of a Metered license.
            Console.WriteLine($"Is metered license accepted: {Aspose.Slides.Metered.IsMeteredLicensed()}");
            //ExEnd:MeteredLicensing
        }

    }
}
