using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace XMindInterviewToDocx
{
    class InterviewDoc : Document
    {
        MainDocumentPart mainDocumentPart;
        bool includeRequirements;

       

        public InterviewDoc(bool includeRequirements)
        {
            
        }

        public MainDocumentPart GenerateMainDocumentPart()
        {
            GenerateMainDocumentPartContent();
            return mainDocumentPart;
        }

        private void GenerateMainDocumentPartContent()
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00C74007", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading1" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Interview Results";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Heading2" };

            paragraphProperties2.Append(paragraphStyleId2);

            Run run2 = new Run();
            Text text2 = new Text();
            text2.Text = "Central Node";

            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "009A1FF9", RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "ListParagraph" };

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = 1 };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "Strong" };

            paragraphMarkRunProperties1.Append(runStyle1);

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(numberingProperties1);
            paragraphProperties3.Append(paragraphMarkRunProperties1);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run3 = new Run() { RsidRunProperties = "009A1FF9" };

            RunProperties runProperties1 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "Strong" };

            runProperties1.Append(runStyle2);
            Text text3 = new Text();
            text3.Text = "Subnode";

            run3.Append(runProperties1);
            run3.Append(text3);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run4 = new Run() { RsidRunProperties = "009A1FF9" };

            RunProperties runProperties2 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "Strong" };

            runProperties2.Append(runStyle3);
            Text text4 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text4.Text = " Level1";

            run4.Append(runProperties2);
            run4.Append(text4);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(proofError1);
            paragraph3.Append(run3);
            paragraph3.Append(proofError2);
            paragraph3.Append(run4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "009A1FF9", RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "ListParagraph" };

            NumberingProperties numberingProperties2 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId2 = new NumberingId() { Val = 1 };

            numberingProperties2.Append(numberingLevelReference2);
            numberingProperties2.Append(numberingId2);

            paragraphProperties4.Append(paragraphStyleId4);
            paragraphProperties4.Append(numberingProperties2);

            Run run5 = new Run();
            Text text5 = new Text();
            text5.Text = "SubnodeLevel2";

            run5.Append(text5);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run5);
            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };
            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };
            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "Heading1" };

            paragraphProperties5.Append(paragraphStyleId5);

            Run run6 = new Run();
            Text text6 = new Text();
            text6.Text = "Requirement Table";

            run6.Append(text6);

            paragraph8.Append(paragraphProperties5);
            paragraph8.Append(run6);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "009A1FF9", RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            Run run7 = new Run();
            Text text7 = new Text();
            text7.Text = "Anything marked with REQUIREMENT before it will be added as a line item to this table.   Add additional columns as needed.";

            run7.Append(text7);

            paragraph9.Append(run7);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "GridTable4-Accent5" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "4675" };
            GridColumn gridColumn2 = new GridColumn() { Width = "4675" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "009A1FF9", RsidTableRowProperties = "0020278B" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            ConditionalFormatStyle conditionalFormatStyle1 = new ConditionalFormatStyle() { Val = "100000000000" };

            tableRowProperties1.Append(conditionalFormatStyle1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            ConditionalFormatStyle conditionalFormatStyle2 = new ConditionalFormatStyle() { Val = "001000000000" };
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(conditionalFormatStyle2);
            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            Run run8 = new Run();
            Text text8 = new Text();
            text8.Text = "Requirement";

            run8.Append(text8);

            paragraph10.Append(run8);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph10);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            ConditionalFormatStyle conditionalFormatStyle3 = new ConditionalFormatStyle() { Val = "100000000000" };

            paragraphProperties6.Append(conditionalFormatStyle3);

            paragraph11.Append(paragraphProperties6);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph11);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "009A1FF9", RsidTableRowProperties = "0020278B" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            ConditionalFormatStyle conditionalFormatStyle4 = new ConditionalFormatStyle() { Val = "000000100000" };

            tableRowProperties2.Append(conditionalFormatStyle4);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            ConditionalFormatStyle conditionalFormatStyle5 = new ConditionalFormatStyle() { Val = "001000000000" };
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(conditionalFormatStyle5);
            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "0020278B" };

            Run run9 = new Run();
            Text text9 = new Text();
            text9.Text = "Req1";

            run9.Append(text9);

            paragraph12.Append(run9);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph12);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties4.Append(tableCellWidth4);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ConditionalFormatStyle conditionalFormatStyle6 = new ConditionalFormatStyle() { Val = "000000100000" };

            paragraphProperties7.Append(conditionalFormatStyle6);

            paragraph13.Append(paragraphProperties7);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph13);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "009A1FF9", RsidTableRowProperties = "0020278B" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            ConditionalFormatStyle conditionalFormatStyle7 = new ConditionalFormatStyle() { Val = "001000000000" };
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(conditionalFormatStyle7);
            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "0020278B" };

            Run run10 = new Run();
            Text text10 = new Text();
            text10.Text = "Req2";

            run10.Append(text10);

            paragraph14.Append(run10);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph14);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            ConditionalFormatStyle conditionalFormatStyle8 = new ConditionalFormatStyle() { Val = "000000000000" };

            paragraphProperties8.Append(conditionalFormatStyle8);

            paragraph15.Append(paragraphProperties8);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph15);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);

            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "0020278B", RsidTableRowProperties = "0020278B" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            ConditionalFormatStyle conditionalFormatStyle9 = new ConditionalFormatStyle() { Val = "000000100000" };

            tableRowProperties3.Append(conditionalFormatStyle9);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            ConditionalFormatStyle conditionalFormatStyle10 = new ConditionalFormatStyle() { Val = "001000000000" };
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties7.Append(conditionalFormatStyle10);
            tableCellProperties7.Append(tableCellWidth7);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "0020278B", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "0020278B" };

            Run run11 = new Run();
            Text text11 = new Text();
            text11.Text = "Req3";

            run11.Append(text11);

            paragraph16.Append(run11);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph16);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "4675", Type = TableWidthUnitValues.Dxa };

            tableCellProperties8.Append(tableCellWidth8);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "0020278B", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "0020278B" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ConditionalFormatStyle conditionalFormatStyle11 = new ConditionalFormatStyle() { Val = "000000100000" };

            paragraphProperties9.Append(conditionalFormatStyle11);

            paragraph17.Append(paragraphProperties9);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph17);

            tableRow4.Append(tableRowProperties3);
            tableRow4.Append(tableCell7);
            tableRow4.Append(tableCell8);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "009A1FF9", RsidParagraphAddition = "009A1FF9", RsidParagraphProperties = "009A1FF9", RsidRunAdditionDefault = "009A1FF9" };
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph18.Append(bookmarkStart1);
            paragraph18.Append(bookmarkEnd1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "009A1FF9", RsidR = "009A1FF9" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(paragraph7);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(table1);
            body1.Append(paragraph18);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart.Document = document1;
        }





    }
}
