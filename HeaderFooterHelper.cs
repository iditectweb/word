using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class HeaderFooterHelper
    {
        public static void AddSimpleHeaderFooter()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));

            //Add header at the left
            Header header = document.Sections[0].Headers.Add();
            Paragraph paragraphHeader = header.Blocks.AddParagraph();
            paragraphHeader.TextAlignment = Styles.Alignment.Left;
            paragraphHeader.Inlines.AddText("simple header");

            //Add footer at the right
            Footer footer = document.Sections[0].Footers.Add();
            Paragraph paragraphFooter = footer.Blocks.AddParagraph();
            paragraphFooter.TextAlignment = Styles.Alignment.Right;
            paragraphFooter.Inlines.AddText("simple footer");

            File.WriteAllBytes("SimpleHeaderFooter.docx", wordFile.Export(document));
        }

        public static void AddHeaderFooterForOddEvenPage()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));
            //Set this property as true to enable odd/even page headers and footers
            document.HasDifferentEvenOddPageHeadersFooters = true;

            //Create odd header with text
            Header headerOdd = document.Sections[0].Headers.Add();
            headerOdd.Blocks.AddParagraph().Inlines.AddText("odd page header");

            //Create even header with text
            Header headerEven = document.Sections[0].Headers.Add(HeaderFooterType.Even);
            headerEven.Blocks.AddParagraph().Inlines.AddText("even page header");

            //Create odd footer with image
            Footer footerOdd = document.Sections[0].Footers.Add(HeaderFooterType.Default);
            using (Stream stream = File.OpenRead("footer1.jpg"))
            {
                footerOdd.Blocks.AddParagraph().Inlines.AddImageInline().Image.ImageSource = new Basic.Media.ImageSource(stream, "jpg");
            }

            //Create even footer with image
            Footer footerEven = document.Sections[0].Footers.Add(HeaderFooterType.Even);
            using (Stream stream = File.OpenRead("footer2.png"))
            {
                footerEven.Blocks.AddParagraph().Inlines.AddImageInline().Image.ImageSource = new Basic.Media.ImageSource(stream, "png");
            }

            File.WriteAllBytes("AddHeaderFooterForOddEvenPage.docx", wordFile.Export(document));
        }

        public static void AddHeaderFooterForSections()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //One section can contains a range of pages

            //Insert one section with single page
            Section sectionSinglePage = builder.InsertSection();
            builder.InsertText("First page in section 1");

            //Add header for single page section
            Header headerSinglePage = sectionSinglePage.Headers.Add();
            headerSinglePage.Blocks.AddParagraph().Inlines.AddText("header for single page section");

            //Insert one section with multiple pages
            Section sectionMultipage = builder.InsertSection();
            //Create first page in section
            builder.InsertText("First page in section 2");
            builder.InsertBreak(BreakType.PageBreak);
            //Create second page in section
            builder.InsertText("Second page in section 2");

            //Defaults, all the secions's header and footer will inherit the rules in the first section
            //If you want to use blank header in the second section, you need initialize a new Header object with nothing to do
            Header headerMultipage = sectionMultipage.Headers.Add();
            //Add footer for multiple page section
            Footer footerMultipage = sectionMultipage.Footers.Add();
            footerMultipage.Blocks.AddParagraph().Inlines.AddText("footer for multiple page section");

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddHeaderFooterForSections.docx"))
            {
                wordFile.Export(document, stream);
            }
        }
    }
}
