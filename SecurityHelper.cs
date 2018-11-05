using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using iDiTect.Word.Protection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class SecurityHelper
    {
        public static void ProtectDocument()
        {
            //Load an existing word file
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Sample.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Protect file with password and permission
            builder.Protect("password", Protection.ProtectionMode.AllowComments);

            File.WriteAllBytes("Protected.docx", wordFile.Export(document));
        }

        public static void UnprotectDocument()
        {
            //Load the protected word file
            WordFile wordFile = new WordFile();
            WordDocument document = wordFile.Import(File.ReadAllBytes("Protected.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //If you have own the password, you can unprotect the document with password
            //builder.Unprotect("password");
            //If you don't own the password, you can unprotect the document without password
            builder.Unprotect();

            File.WriteAllBytes("Unprotected.docx", wordFile.Export(document));
        }

        
    }
}
