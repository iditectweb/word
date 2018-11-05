using iDiTect.Word.Basic.Model;
using iDiTect.Word.Editing;
using iDiTect.Word.IO;
using iDiTect.Word.Shapes;
using iDiTect.Word.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;

namespace iDiTect.Word.Demo
{
    public static class TableHelper
    {
        public static void AddSimpleTable()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);            
            builder.TableState.Indent = 100;
            
            Paragraph tableTitle = builder.InsertParagraph();
            tableTitle.TextAlignment = Alignment.Center;
            builder.InsertLine("Simple Table Title");

            Table table = builder.InsertTable();
            table = CreateSimpleTable(table); 

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddSimpleTable.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        public static void AddTableFrame()
        {
            WordDocument document = new WordDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);
            builder.TableState.Indent = 100;
            builder.TableState.PreferredWidth = new TableWidthUnit(300);
            builder.TableState.LayoutType = TableLayoutType.FixedWidth;

            Paragraph tableTitle = builder.InsertParagraph();
            tableTitle.TextAlignment = Alignment.Center;
            builder.InsertLine("Table Frame");

            Table table = builder.InsertTable();
            table = CreateTableFrame(table);

            WordFile wordFile = new WordFile();
            using (var stream = File.OpenWrite("AddTableFrame.docx"))
            {
                wordFile.Export(document, stream);
            }
        }

        public static Table CreateSimpleTable(Table table)
        {           
            ThemableColor bordersColor =new ThemableColor(Color.FromRgb(73, 90, 128));
            ThemableColor headerColor = new ThemableColor(Color.FromRgb(34, 143, 189));
            ThemableColor defaultRowColor = new ThemableColor(Color.FromRgb(176, 224, 230));

            Border border = new Border(1, BorderStyle.Single, bordersColor);
            table.Borders = new TableBorders(border);
            table.TableCellPadding = new Basic.Primitives.Padding(5);
                        
            //Add table header
            TableRow headerRow = table.Rows.AddTableRow();
            headerRow.RepeatOnEveryPage = true;
            //Add first column
            TableCell column1 = headerRow.Cells.AddTableCell();
            column1.State.BackgroundColor.LocalValue = headerColor;
            column1.Borders = new TableCellBorders(border, border, border, border, null, null, border, null);
            column1.PreferredWidth = new TableWidthUnit(50);
            //Add second column
            TableCell column2 = headerRow.Cells.AddTableCell();
            column2.State.BackgroundColor.LocalValue = headerColor;
            column2.PreferredWidth = new TableWidthUnit(150);
            column2.VerticalAlignment = VerticalAlignment.Center;
            Paragraph column2Para = column2.Blocks.AddParagraph();
            column2Para.TextAlignment = Alignment.Center;
            column2Para.State.LineSpacing = 1;
            TextInline column2Text = column2Para.Inlines.AddText("Product");
            column2Text.State.ForegroundColor = new ThemableColor(Colors.White);
            column2Text.FontSize = 20;
            //Add third column
            TableCell column3 = headerRow.Cells.AddTableCell();
            column3.State.BackgroundColor.LocalValue = headerColor;
            column3.PreferredWidth = new TableWidthUnit(250);
            column3.Padding = new Basic.Primitives.Padding(20, 0, 0, 0);
            Paragraph column3Para = column3.Blocks.AddParagraph();
            column3Para.State.LineSpacing = 1;
            TextInline column3Text = column3Para.Inlines.AddText("Price");
            column3Text.State.ForegroundColor = new ThemableColor(Colors.White);
            column3Text.FontSize = 20;
            
            //Add table rows
            Random r = new Random();
            for (int i = 0; i < 50; i++)
            {
                ThemableColor rowColor = i % 2 == 0 ? defaultRowColor : new ThemableColor(Colors.White);

                TableRow row = table.Rows.AddTableRow();
                row.Height = new TableRowHeight(HeightType.Exact, 20);

                TableCell idCell = row.Cells.AddTableCell();                
                idCell.State.BackgroundColor.LocalValue = rowColor;
                idCell.Blocks.AddParagraph().Inlines.AddText(i.ToString());

                TableCell productCell = row.Cells.AddTableCell();
                productCell.State.BackgroundColor.LocalValue = rowColor;
                Paragraph productPara = productCell.Blocks.AddParagraph();
                productPara.TextAlignment = Alignment.Center;
                productPara.Inlines.AddText(String.Format("Product{0}", i));

                TableCell priceCell = row.Cells.AddTableCell();
                priceCell.Padding = new Basic.Primitives.Padding(20, 0, 0, 0);
                priceCell.State.BackgroundColor.LocalValue = rowColor;
                priceCell.Blocks.AddParagraph().Inlines.AddText(r.Next(10, 1000).ToString());
            }

            return table;
        }

        public static Table CreateTableFrame(Table table)
        {
            ThemableColor bordersColor =new ThemableColor(Colors.Black);

            //Set table border
            table.Borders = new TableBorders(new Border(3, BorderStyle.Single, bordersColor));
            table.TableCellPadding = new Basic.Primitives.Padding(6);
                        
            TableRow row = table.Rows.AddTableRow();

            //Add a merged cell in 2x2
            TableCell cell = row.Cells.AddTableCell();
            cell.RowSpan = 2;
            cell.ColumnSpan = 2;
            cell.Blocks.AddParagraph().Inlines.AddText("Text 1");

            //Add a single cell
            cell = row.Cells.AddTableCell();
            cell.Blocks.AddParagraph().Inlines.AddText("Text 2");

            row = table.Rows.AddTableRow();
            //Add a single cell
            cell = row.Cells.AddTableCell();
            cell.Blocks.AddParagraph().Inlines.AddText("Text 3");

            row = table.Rows.AddTableRow();
            //Add a single cell
            cell = row.Cells.AddTableCell();
            cell.Blocks.AddParagraph().Inlines.AddText("Text 4");

            //Add a merged cell in 1x2
            cell = row.Cells.AddTableCell();
            cell.ColumnSpan = 2;
            cell.Blocks.AddParagraph().Inlines.AddText("Text 5");

            row = table.Rows.AddTableRow();
            //Add a single cell
            cell = row.Cells.AddTableCell();
            cell.Blocks.AddParagraph().Inlines.AddText("Text 6");

            //Add a single cell
            cell = row.Cells.AddTableCell();
            cell.Blocks.AddParagraph().Inlines.AddText("Text 7");
           
            //Add a single cell
            cell = row.Cells.AddTableCell();
            ImageInline imageCell = cell.Blocks.AddParagraph().Inlines.AddImageInline();
            using (Stream stream = File.OpenRead("watermark.png"))
            {
                imageCell.Image.ImageSource = new Basic.Media.ImageSource(stream, "png");
                imageCell.Image.Width = 50;
                imageCell.Image.Height = 50;
            }

            return table;
        }



    }
}
