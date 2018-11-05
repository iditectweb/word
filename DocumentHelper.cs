using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class DocumentHelper
    {
        public static void MergeDocument()        
        {
            WordFile wordFile = new WordFile();
            WordDocument source = wordFile.Import(File.ReadAllBytes("source.docx"));
            WordDocument target = wordFile.Import(File.ReadAllBytes("target.docx"));

            //The source document will be appended at the end of target document
            //The content of source document will be started at a new page
            target.Merge(source);

            //If these two documents have the same style ID, you can choose which to use
            //MergeOptions options = new MergeOptions();
            //options.ConflictingStylesResolutionMode = ConflictingStylesResolutionMode.UseTargetStyle;
            //target.Merge(source, options);

            File.WriteAllBytes("Merged.docx", wordFile.Export(target));
        }

        public static void InsertDocument()
        {
            WordFile wordFile = new WordFile();
            WordDocument source = wordFile.Import(File.ReadAllBytes("source.docx"));
            WordDocument target = new WordDocument();

            WordDocumentBuilder builder = new WordDocumentBuilder(target);
            builder.CharacterState.FontSize = 30;

            builder.InsertLine("Text start in target document.");           

            //Insert the source document just after the first paragraph in the target document
            builder.InsertDocument(source);

            //This line will be appended tight after the source document content
            builder.InsertLine("Text end in target document.");

            File.WriteAllBytes("InsertDocument.docx", wordFile.Export(target));
        }
    }
}
