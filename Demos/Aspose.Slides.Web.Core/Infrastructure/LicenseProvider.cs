using System;
using System.IO;
using Aspose.Slides.Web.Core.Enums;

namespace Aspose.Slides.Web.Core.Infrastructure
{
	///<Summary>
	/// License class to set Aspose products license
	///</Summary>
	internal sealed class LicenseProvider : ILicenseProvider
	{
		public string LicensePath { get; }

		public LicenseProvider(string licensePath)
		{
			LicensePath = licensePath;
		}

		///<Summary>
		/// SetAsposeLicense method to setup license for Aspose product
		///</Summary>
		public void SetAsposeLicense(AsposeProducts asposeProducts)
		{
			if (!File.Exists(LicensePath))
			{
				return;
			}

			switch (asposeProducts)
			{
				case AsposeProducts.Cells:
					{
						var acLic = new Cells.License();
						acLic.SetLicense(LicensePath);

						break;
					}

				case AsposeProducts.Slides:
					{
						var acLic = new Slides.License();
						acLic.SetLicense(LicensePath);

						break;
					}

				case AsposeProducts.Words:
					{
						var acLic = new Words.License();
						acLic.SetLicense(LicensePath);

						break;
					}

				default:
					{
						throw new NotSupportedException($"The type: {asposeProducts} of the AsposeProducts doesn't support."); 
					}
			}
		}
	}
}
