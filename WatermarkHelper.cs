using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using iDiTect.Word.Watermarks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;

namespace iDiTect.Word.Demo
{
    public static class WatermarkHelper
    {
        public static void AddTextWatermark()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));

            //Customize the setting of text watermark
            TextWatermarkSettings setting = new TextWatermarkSettings();
            setting.Text = "watermark";
            setting.Width = 100;
            setting.Height = 50;
            //The opacity value is between 0 and 1.
            setting.Opacity = 0.7;
            //Set the watermark rotation
            setting.Rotation = -45;
            setting.TextColor = Colors.Red;

            //Create watermark with settings
            Watermark textWatermark = new Watermark(setting);

            //Add watermark to Header object
            Header header = document.Sections[0].Headers.Add();
            header.Watermarks.Add(textWatermark);
                       
            File.WriteAllBytes("AddTextWatermark.docx", wordFile.Export(document));
        }

        public static void AddImageWatermark()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Customize the setting of image watermark
            ImageWatermarkSettings setting = new ImageWatermarkSettings();
            setting.Width = 100;
            setting.Height = 50;
            setting.Rotation = -45;
            using (Stream stream = File.OpenRead("watermark.png"))
            {
                setting.ImageSource = new Basic.Media.ImageSource(stream, "png");
            }

            //Create watermark with settings
            Watermark imageWatermark = new Watermark(setting);

            //Add watermark to Header object
            builder.SetWatermark(imageWatermark, document.Sections[0].Headers.Add());
            //builder.SetWatermark(imageWatermark, document.Sections[0], HeaderFooterType.Default);

            File.WriteAllBytes("AddImageWatermark.docx", wordFile.Export(document));
        }
    }
}
