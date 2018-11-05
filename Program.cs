using iDiTect.Word.Basic.Media;
using iDiTect.Word.Editing;
using iDiTect.Word.Fields;
using iDiTect.Word.IO;
using iDiTect.Word.Licensing;
using iDiTect.Word.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            //This license registration line need to be at very beginning of our other code
            LicenseManager.SetKey("CLBUM-YGHWC-TCJFY-R3QJ5-3PRZZ-HCMA6");

            //TextHelper.AddText();
            //ImageHelper.AddInlineImage();
            //HeaderFooterHelper.AddSimpleHeaderFooter();
            //BookmarkHelper.AddBookmark();
            //CommentHelper.AddComment();
            //SecurityHelper.ProtectDocument();
            //LinkHelper.AddLinkInsideDocument();
            //WatermarkHelper.AddImageWatermark();
            //DocumentHelper.MergeDocument();
            //TableHelper.AddSimpleTable();
            MailMergeHelper.AddMailMerge();

           
        }
    }
}
