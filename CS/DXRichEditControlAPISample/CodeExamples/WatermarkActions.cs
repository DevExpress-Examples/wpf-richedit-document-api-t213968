using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

namespace DXRichEditControlAPISample.CodeExamples
{
    class WatermarkActions
    {
        static void CreateTextWatermark(Document document)
        {
            #region #CreateTextWatermark
            //Check whether the document sections have headers:
            foreach (Section section in document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    //If not, create an empty header
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.Color = System.Drawing.Color.LightGray;
            textWatermarkOptions.FontFamily = "Calibri";
            textWatermarkOptions.Layout = WatermarkLayout.Horizontal;
            textWatermarkOptions.Semitransparent = true;

            document.WatermarkManager.SetText("CONFIDENTIAL", textWatermarkOptions);
            #endregion #CreateTextWatermark
        }
        static void CreateImageWatermark(Document document)
        {
            #region #CreateImageWatermark
            //Check whether the document sections have headers:
            foreach (Section section in document.Sections)
            {
                if (!section.HasHeader(HeaderFooterType.Primary))
                {
                    //If not, create an empty header
                    SubDocument header = section.BeginUpdateHeader();
                    section.EndUpdateHeader(header);
                }
            }

            ImageWatermarkOptions imageWatermarkOptions = new ImageWatermarkOptions();
            imageWatermarkOptions.Washout = false;
            imageWatermarkOptions.Scale = 2;
            document.WatermarkManager.SetImage(System.Drawing.Image.FromFile("Documents//DevExpress.png"), imageWatermarkOptions);
            #endregion #CreateImageWatermark

        }

    }
}
