using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using iDiTect.Word.Shapes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class ImageHelper
    {        
        public static void AddInlineImage()
        {            
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Add in-line image using builder
            builder.InsertText("Simple sentence 1 in line. ");
            using (Stream stream = File.OpenRead("sample.jpg"))
            {
                builder.InsertImageInline(stream, "jpg");
            }
            builder.InsertText("Simple sentence 2 in line");

            //Add in-line image using Paragraph object
            Paragraph paragraph = builder.InsertParagraph();
            //Add text before
            TextInline textStart = paragraph.Inlines.AddText();
            textStart.Text = "Text add using paragraph start.";
            //Insert image in the middle of text content
            ImageInline imageInline = paragraph.Inlines.AddImageInline();
            using (Stream stream = File.OpenRead("sample.png"))
            {
                imageInline.Image.ImageSource = new Basic.Media.ImageSource(stream, "png");
            }
            //Add text after
            TextInline textEnd = paragraph.Inlines.AddText();
            textEnd.Text = "Text add using paragraph end.";

            WordFile wordFile = new WordFile();
            File.WriteAllBytes("AddImageInline.docx", wordFile.Export(document));
        }

        public static void AddFloatingImage()
        {            
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);
            builder.CharacterState.FontSize = 24;

            //Add floating image using builder            
            using (Stream stream = File.OpenRead("sample.jpg"))
            {
                FloatingImage floatingImage1 = builder.InsertFloatingImage(stream, "jpg");
                floatingImage1.Wrapping.WrappingType = ShapeWrappingType.Square;
            }
            builder.InsertText("This text sentence content will display at the square of the floating image. ");
            builder.InsertText("This text sentence content will display at the square of the floating image.");

            WordFile wordFile = new WordFile();
            File.WriteAllBytes("AddFloatingImage.docx", wordFile.Export(document));
        }
    }
}
