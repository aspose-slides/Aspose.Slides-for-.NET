using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Config;
using System.Web;
using System.Globalization;
using System.Text;
using Aspose.Slides.Live.Demos.UI.Helpers;
namespace Aspose.Slides.Live.Demos.UI
{
  public partial class WatermarkSlides : BasePage
  {
		AsposeSlides asposeSlides = new AsposeSlides();
		protected void Page_Load(object sender, EventArgs e)
    {
      Product = "slides";
      Page.Title = Resources[Product + "WatermarkPageTitle"];
      Page.MetaDescription = Resources[Product + "WatermarkMetaDescription"];
	  AsposeProductTitle.InnerText = PageProductTitle + " " + Resources["WatermarkAPPName"];
      ProductTitle.InnerText = Resources[Product + "WatermarkTitle"];
      ProductTitleSub.InnerText = Resources[Product + "WatermarkSubTitle"];
      ProductImage.Src = "~/img/aspose-" + Product + "-app.png";
      PoweredBy.InnerText = PageProductTitle + ". ";
      PoweredBy.HRef = "https://products.aspose.com/" + Product;
      TextWatermarkButton.Text = Resources["TextWatermarkButton"];
      ProcessTextWatermarkButton.Text = Resources["TextWatermarkButton"];
      ImageWatermarkButton.Text = Resources["ImageWatermarkButton"];
      ProcessImageWatermarkButton.Text = Resources["ImageWatermarkButton"];
      RemoveWatermarkButton.Text = Resources["RemoveWatermarkButton"];
      textWatermark.Attributes.Add("placeholder", Resources["AddWatermarkTextPlaceholder"]);
      
			ViewState["AddanotherWatermark"] = HttpContext.Current.Request.Url.AbsoluteUri;
			ViewState["product"] = Product;
			// Check for auto-generate URLs to set only format as valid extension
			if (Page.RouteData.Values["Format"] != null)
			{
				ViewState["validFileExtensions"] = SetValidation("." + Page.RouteData.Values["Format"].ToString().ToLower(), ValidateFileType);
			}
			else
			{
				ViewState["validFileExtensions"] = SetValidation(Resources[Product + "WatermarkValidationExpression"], ValidateFileType);
			}
      SetValidation(Resources[Product + "WatermarkImageValidationExpression"], ValidateImageType);
     
			
		}

    protected void UploadFile(Action<FileUploadResponse> action, string folderName)
    {
      if (IsValid)
        if (CheckFileInputs(UploadFileInput))
          try
          {
            var files = UploadFiles( folderName, UploadFileInput);
            if (files != null && files.Count == 1)
            {
              FileNameHidden.Value = files[0].FileName;
              FolderNameHidden.Value = files[0].FolderId;
              action(files[0]);
            }
          }
          catch (Exception ex)
          {
            ShowErrorMessage(WatermarkMessage, "Error: " + ex.Message);
          }
        else
          ShowErrorMessage(WatermarkMessage, Resources["FileSelectMessage"]);
    }

    protected void TextWatermarkButton_Click(object sender, EventArgs e)
    {
      UploadFile((file)  =>
      {
        UploadFilePlaceHolder.Visible = false;
        TextPlaceHolder.Visible = true;
        ProductTitleSub.InnerText = Resources[Product + "WatermarkTextTitleSub"];
      }, Guid.NewGuid().ToString());
    }

    protected void ProcessTextWatermarkButton_Click(object sender, EventArgs e)
    {

			TextWatermarkOptionsModel textWatermarkOptionsModel = new TextWatermarkOptionsModel();
			textWatermarkOptionsModel.id = FolderNameHidden.Value ;
			textWatermarkOptionsModel.FileName = FileNameHidden.Value;

			textWatermarkOptionsModel.Text = textWatermark.Value;
			textWatermarkOptionsModel.Color = pickcolor.Value;
			textWatermarkOptionsModel.FontName = fontFamily.SelectedValue;
			textWatermarkOptionsModel.FontSize = int.Parse(fontSize.Text);
			textWatermarkOptionsModel.RotationAngleDegrees = int.Parse(textAngle.Text);

			var response = asposeSlides.AddTextWatermark(textWatermarkOptionsModel);
			SuccessLabel.InnerText = Resources["WatermarkAddedSuccessMessage"];
			PerformResponse(response, TextMessage, ShowDownloadPage);
		}

    protected void ImageWatermarkButton_Click(object sender, EventArgs e)
    {
      UploadFile((file) =>
      {
        UploadFilePlaceHolder.Visible = false;
        ImagePlaceHolder.Visible = true;
        ProductTitleSub.InnerText = Resources[Product + "WatermarkImageTitleSub"];
      }, Guid.NewGuid().ToString());
    }

    protected void ProcessImageWatermarkButton_Click(object sender, EventArgs e)
    {
			if (IsValid)
				if (CheckFileInputs(UploadImageInput))
					try
					{
						var files = UploadFiles(FolderNameHidden.Value, UploadImageInput);
						if (files != null && files.Count == 1)
						{

							ImageWatermarkOptionsModel imageWatermarkOptionsModel = new ImageWatermarkOptionsModel();
							imageWatermarkOptionsModel.id = FolderNameHidden.Value;
							imageWatermarkOptionsModel.FileName = files[0].FileName ;

							imageWatermarkOptionsModel.idMain = FolderNameHidden.Value;
							imageWatermarkOptionsModel.MainFileName = FileNameHidden.Value;

							imageWatermarkOptionsModel.IsGrayScaled = greyScale.Checked;
							imageWatermarkOptionsModel.ZoomPercent = int.Parse(zoom.Text);
							imageWatermarkOptionsModel.RotationAngleDegrees = int.Parse(imageAngle.Text);


							//var response = asposeSlides.AddImageWatermark(FileNameHidden.Value, FolderNameHidden.Value, files[0].FileName, files[0].FolderId, greyScale.Checked, double.Parse(zoom.Text), double.Parse(imageAngle.Text));
							var response = asposeSlides.AddImageWatermark(imageWatermarkOptionsModel);
							SuccessLabel.InnerText = Resources["WatermarkAddedSuccessMessage"];
							PerformResponse(response, ImageMessage, ShowDownloadPage);
						}
					}
					catch (Exception ex)
					{
						ShowErrorMessage(ImageMessage, "Error: " + ex.Message);
					}
				else
					ShowErrorMessage(ImageMessage, Resources["FileSelectMessage"]);
		}

    protected void RemoveWatermarkButton_Click(object sender, EventArgs e)
    {
			UploadFile((file) =>
			{


				BaseRequestModel baseRequestModel = new BaseRequestModel();
				baseRequestModel.id = FolderNameHidden.Value ;
				baseRequestModel.FileName = FileNameHidden.Value;

				

				var response =  asposeSlides.RemoveWatermark(baseRequestModel);
				SuccessLabel.InnerText = Resources["WatermarkRemovedSuccessMessage"];
				PerformResponse(response, WatermarkMessage, ShowDownloadPage);
			}, Guid.NewGuid().ToString());
		}

    private void ShowDownloadPage(FileSafeResult response)
    {
			var url = response.DownloadURL();
			var callbackURL = HttpContext.Current.Request.Url.AbsolutePath;
			//var viewerURL = response.ViewerURL(Product, callbackURL);
			DownloadButton.NavigateUrl = url;
			// DownloadUrlInputHidden.Value = HttpUtility.UrlEncode(url);
			//ViewerLink.NavigateUrl = viewerURL;
			UploadFilePlaceHolder.Visible = false;
			TextPlaceHolder.Visible = false;
			ImagePlaceHolder.Visible = false;
			DownloadPlaceHolder.Visible = true;
		}

    
  }
}
