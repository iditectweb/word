using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace iDiTect.Word.Demo
{
    public static class MailMergeHelper
    {
        public static void AddMailMerge()
        {
            WordDocument document = CreateMailMergeTemplate();
            IEnumerable datas = CreateTestData();
            WordDocument mailMergedDocument = document.MailMerge(datas);
            

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("MailMergeData.docx"))
            {
                wordFile.Export(mailMergedDocument, stream);
            }
        }

        public static WordDocument CreateMailMergeTemplate()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);

            //Insert salutation
            builder.InsertText("Hello ");
            builder.InsertField("MERGEFIELD CustomerFirstName", "");
            builder.InsertText(" ");
            builder.InsertField("MERGEFIELD CustomerLastName", "");
            builder.InsertText(",");

            //Insert a blank line
            builder.InsertParagraph();

            //Insert mail body
            builder.InsertParagraph();
            builder.InsertText("Thanks for purchasing our ");
            builder.InsertField("MERGEFIELD ProductName ", "");
            builder.InsertText(", please download your Invoice at ");
            builder.InsertField("MERGEFIELD InvoiceURL", "");
            builder.InsertText(". If you have any questions please call ");
            builder.InsertField("MERGEFIELD SupportFhone", "");
            builder.InsertText(", or email us at ");
            builder.InsertField("MERGEFIELD SupportEmail", "");
            builder.InsertText(".");

            //Insert a blank line
            builder.InsertParagraph();

            //Insert mail ending
            builder.InsertParagraph();
            builder.InsertText("Best regards,");
            builder.InsertBreak(BreakType.LineBreak);
            builder.InsertField("MERGEFIELD EmployeeFullname", "");
            builder.InsertText(" ");
            builder.InsertField("MERGEFIELD EmployeeDepartment", "");

            return document;
        }

        public static IEnumerable CreateTestData()
        {
            List<MailMergeObject> datas = new List<MailMergeObject>();

            var data = new MailMergeObject
            {
                CustomerFirstName = "test-customer-first-name-1",
                CustomerLastName = "test-customer-last-name-1",
                ProductName = "test-product-name-1",
                InvoiceURL = "test-invoice-url-1",
                SupportFhone = "test-support-phone-1",
                SupportEmail = "test-support-email-1",
                EmployeeFullname = "test-employee-fullname-1",
                EmployeeDepartment = "test-employee-department-1"
            };

            datas.Add(data);

            data = new MailMergeObject
            {
                CustomerFirstName = "test-customer-first-name-2",
                CustomerLastName = "test-customer-last-name-2",
                ProductName = "test-product-name-2",
                InvoiceURL = "test-invoice-url-2",
                SupportFhone = "test-support-phone-2",
                SupportEmail = "test-support-email-2",
                EmployeeFullname = "test-employee-fullname-2",
                EmployeeDepartment = "test-employee-department-2"
            };

            datas.Add(data);
            return datas;
        }
    }

    public class MailMergeObject
    {
        public string CustomerFirstName { get; set; }
        public string CustomerLastName { get; set; }
        public string ProductName { get; set; }
        public string InvoiceURL { get; set; }
        public string SupportFhone { get; set; }
        public string SupportEmail { get; set; }
        public string EmployeeFullname { get; set; }
        public string EmployeeDepartment { get; set; }
    }
}
