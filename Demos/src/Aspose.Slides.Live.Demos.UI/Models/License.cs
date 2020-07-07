using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	///<Summary>
	/// License class to set apose products license
	///</Summary>
	public static class License
	{
		private static string _licenseFileName = "Aspose.Total.lic";


		///<Summary>
		/// SetAsposeSlidesLicense method to Aspose.Words License
		///</Summary>
		public static void SetAsposeSlidesLicense()
		{
			try
			{
				Aspose.Slides.License acLic = new Aspose.Slides.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		public static void SetAsposeWordsLicense()
		{
			try
			{
				Aspose.Words.License acLic = new Aspose.Words.License();
				acLic.SetLicense(_licenseFileName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
	}
}
