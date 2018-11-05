using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class CommentHelper
    {
        public static void AddComment()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            builder.InsertLine("First paragraph in this document.");

            //Insert some content before comment
            builder.InsertText("Second paragraph start. ");
            //Select the content you want to comment
            TextInline textComment = builder.InsertText("Text has comment. ");
            //Insert some content after comment
            builder.InsertText("Second paragraph end.");

            //Add comment with selected content
            Comment comment = builder.InsertComment("Comment details here", textComment, textComment);
            comment.Author = "iDiTect";
            comment.Date = DateTime.Now;

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddComment.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        public static void AddComment2()
        {
            WordDocument document = new WordDocument();

            //Insert some content before bookmark
            Section section = document.Sections.AddSection();
            Paragraph para = section.Blocks.AddParagraph();
            para.Inlines.AddText("Sentence start ");

            //Create comment
            Comment comment = document.Comments.AddComment();
            comment.Author = "iDiTect";
            comment.Date = DateTime.Now;
            comment.Blocks.AddParagraph().Inlines.AddText("comment details");

            //Insert comment to paragraph
            para.Inlines.Add(comment.CommentRangeStart);
            para.Inlines.AddText("text");
            para.Inlines.Add(comment.CommentRangeEnd);

            //Insert some content after bookmark
            para.Inlines.AddText(" Sentence end.");

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddComment2.docx"))
            {
                wordFile.Export(document, stream);
            }
        }
    }
}
