using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class LinkHelper
    {
        public static void AddLinkToWebLink()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Add hyperlink to web url
            builder.InsertHyperlink("this is hyperlink", "http://www.iditect.com", "go to iditect site");

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddLinkToWebLink.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        public static void AddLinkInsideDocument()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Add hyperlink navigate to bookmark inside this document by bookmark's name
            builder.InsertHyperlinkToBookmark("this is hyerlink", "bookmark1", "go to bookmark1");

            //Add a bookmark in the second page
            builder.InsertBreak(BreakType.PageBreak);
            TextInline textBookmark = builder.InsertText("This is bookmark1. ");
            builder.InsertBookmark("bookmark1", textBookmark, textBookmark);

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddLinkInsideDocument.docx"))
            {
                wordFile.Export(document, stream);
            }
        }
    }
}
