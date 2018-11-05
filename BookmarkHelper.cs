using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class BookmarkHelper
    {
        public static void AddBookmark()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);
                        
            builder.InsertLine("First paragraph in this document.");

            //Insert some content before bookmark
            builder.InsertText("Second paragraph start. ");
            //Select the content you want to bookmark
            TextInline textBookmark = builder.InsertText("This is bookmark. ");
            //Insert some content after bookmark
            builder.InsertText("Second paragraph end.");

            //Add bookmark with selected content
            builder.InsertBookmark("bookmark1", textBookmark, textBookmark);
                      
            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddBookmark.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        public static void AddBookmark2()
        {
            WordDocument document = new WordDocument();

            //Insert some content before bookmark
            Section section = document.Sections.AddSection();
            Paragraph para = section.Blocks.AddParagraph();
            para.Inlines.AddText("Sentence start ");

            //Create bookmark
            Bookmark bookmark = new Bookmark(document, "bookmark2");
            para.Inlines.Add(bookmark.BookmarkRangeStart);
            para.Inlines.AddText("text");
            para.Inlines.Add(bookmark.BookmarkRangeEnd);

            //Insert some content after bookmark
            para.Inlines.AddText(" Sentence end.");

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddBookmark2.docx"))
            {
                wordFile.Export(document, stream);
            }
        }
    }
}
