using iDiTect.Word.Basic.Model;
using iDiTect.Word.Basic.Theming;
using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using iDiTect.Word.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace iDiTect.Word.Demo
{
    public static class TextHelper
    {
        public static void AddText()
        {
            WordDocument document = new WordDocument();            
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Set global style for text and paragraph
            builder.CharacterState.FontFamily = new ThemableFontFamily("Arial");
            builder.CharacterState.FontSize = 16;
            builder.ParagraphState.LineSpacing = 1.2;
            builder.ParagraphState.FirstLineIndent = 40;

            //Insert text using builder directly            
            builder.InsertText("Nomal text. ");
            //Insert one line with text, it will add line break automatically
            builder.InsertLine("Nomal line with auto line break. ");
            //So the text below will be added in a second paragraph
            builder.InsertText("Nomal text. ");

            //Insert text using TextInline object
            TextInline textInline = new TextInline(document);
            textInline.Text = "This text content is using TextInline object. ";
            textInline.FontSize = 20;
            builder.InsertInline(textInline);

            //Insert text with customized style
            builder.InsertText("Times New Roman, ").FontFamily = new ThemableFontFamily("Times New Roman");
            builder.InsertText("bold, ").FontWeight = FontWeights.Bold;
            builder.InsertText("italic, ").FontStyle = FontStyles.Italic;
            builder.InsertText("underline, ").Underline.Pattern = UnderlinePattern.Single;
            builder.InsertText("colors ").ForegroundColor = new ThemableColor(Color.FromRgb(255, 0, 0));

            //Add several paragraphs to page           
            for (int i = 0; i < 20; i++)
            {
                builder.InsertParagraph();
                for (int j = 1; j < 11; j++)
                {
                    builder.InsertText("This is sentence " + j.ToString() + ". ");
                }
            }     

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddText.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        
        public static void ReplaceText()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));

            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Replace target text in whole document, match-case and match-whole-word are supported
            builder.ReplaceText("Page", "as", true, true);

            File.WriteAllBytes("ReplaceText.docx", wordFile.Export(document));
        }

        public static void HighlightText()
        {
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));

            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Apply new highlight style 
            Action<CharacterState> action = new Action<CharacterState>((state) =>
            {
                state.HighlightColor = Colors.Yellow;
            });

            //Highlight all the "Page" text in the document
            builder.ReplaceStyling("Page", true, true, action);

            File.WriteAllBytes("HighlightText.docx", wordFile.Export(document));
        }
    }
}
