using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using System.Collections.Generic;
using System;
using System.IO;

namespace ReportPublish
{
    public struct ThongTinGiangDay
    {
        public string MonHoc;
        public string Lop;
        public string BacDaoTao;
        public double GioChuan;

        public ThongTinGiangDay(string monHoc, string lop, string bacDaoTao, double gioChuan)
        {
            MonHoc = monHoc;
            Lop = lop;
            BacDaoTao = bacDaoTao;
            GioChuan = gioChuan;
        }
    }

    public struct GiaoVien
    {
        public string Ten;
        public string MSCB;
        public List<ThongTinGiangDay> Ttgd;

        public GiaoVien(string ten, string mSCB)
        {
            Ten = ten;
            MSCB = mSCB;
            Ttgd = new List<ThongTinGiangDay>();
        }
    }

    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        /*
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }
        */

        public void CreatePackage(string filePath, GiaoVien gv)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package, gv);
            }
        }
        public byte[] CreatePackage(GiaoVien gv)
        {
            using (var ms = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                    CreateParts(package, gv);
                return ms.ToArray();
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document, GiaoVien gv)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1, gv);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId5");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId4");
            GenerateFontTablePart1Content(fontTablePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "18";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "165";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "945";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "7";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "2";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Title";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1108";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1, GiaoVien gv)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
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
            document1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            document1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            document1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "10978", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableJustification1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "5353" };
            GridColumn gridColumn2 = new GridColumn() { Width = "5625" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00B42C64", RsidTableRowAddition = "004632DA", RsidTableRowProperties = "00FF498E", ParagraphId = "63C289C6", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties1.Append(tableJustification2);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "5353", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "004632DA", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "004632DA", ParagraphId = "687741DE", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation1 = new Indentation() { End = "-357" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold1 = new Bold();

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);

            paragraphProperties1.Append(spacingBetweenLines1);
            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };

            runProperties1.Append(runFonts2);
            Text text1 = new Text();
            text1.Text = "ĐẠI HỌC QUỐC GIA TP.HCM";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "5625", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "004632DA", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "004632DA", ParagraphId = "4226C0E3", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation2 = new Indentation() { End = "-357" };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold2 = new Bold();

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(bold2);

            paragraphProperties2.Append(spacingBetweenLines2);
            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold3 = new Bold();

            runProperties2.Append(runFonts4);
            runProperties2.Append(bold3);
            Text text2 = new Text();
            text2.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00B42C64", RsidTableRowAddition = "004632DA", RsidTableRowProperties = "00FF498E", ParagraphId = "6F436B90", TextId = "77777777" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableJustification tableJustification3 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties2.Append(tableJustification3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "5353", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "004632DA", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "004632DA", ParagraphId = "0D891744", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation3 = new Indentation() { End = "-357" };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold4 = new Bold();

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(bold4);

            paragraphProperties3.Append(spacingBetweenLines3);
            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold5 = new Bold();

            runProperties3.Append(runFonts6);
            runProperties3.Append(bold5);
            Text text3 = new Text();
            text3.Text = "TR";

            run3.Append(runProperties3);
            run3.Append(text3);

            Run run4 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.ComplexScript, Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold6 = new Bold();

            runProperties4.Append(runFonts7);
            runProperties4.Append(bold6);
            Text text4 = new Text();
            text4.Text = "Ư";

            run4.Append(runProperties4);
            run4.Append(text4);

            Run run5 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold7 = new Bold();

            runProperties5.Append(runFonts8);
            runProperties5.Append(bold7);
            Text text5 = new Text();
            text5.Text = "ỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);
            paragraph3.Append(run4);
            paragraph3.Append(run5);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "00906534", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00906534", ParagraphId = "392C4B72", TextId = "3B77C17C" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation4 = new Indentation() { End = "-357" };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold8 = new Bold();

            paragraphMarkRunProperties4.Append(runFonts9);
            paragraphMarkRunProperties4.Append(bold8);

            paragraphProperties4.Append(spacingBetweenLines4);
            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold9 = new Bold();

            runProperties6.Append(runFonts10);
            runProperties6.Append(bold9);
            Text text6 = new Text();
            text6.Text = "--------------------------------------";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run6);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);
            tableCell3.Append(paragraph4);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "5625", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "004632DA", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "004632DA", ParagraphId = "0533DAE2", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation5 = new Indentation() { End = "-357" };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold10 = new Bold();
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties5.Append(runFonts11);
            paragraphMarkRunProperties5.Append(bold10);
            paragraphMarkRunProperties5.Append(fontSize1);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript1);

            paragraphProperties5.Append(spacingBetweenLines5);
            paragraphProperties5.Append(indentation5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run7 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold11 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "26" };

            runProperties7.Append(runFonts12);
            runProperties7.Append(bold11);
            runProperties7.Append(fontSize2);
            runProperties7.Append(fontSizeComplexScript2);
            Text text7 = new Text();
            text7.Text = "Độc lập – Tự do – Hạnh phúc";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run7);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "00906534", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00906534", ParagraphId = "73571812", TextId = "1D178DF4" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation6 = new Indentation() { End = "-357" };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold12 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties6.Append(runFonts13);
            paragraphMarkRunProperties6.Append(bold12);
            paragraphMarkRunProperties6.Append(fontSize3);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript3);

            paragraphProperties6.Append(spacingBetweenLines6);
            paragraphProperties6.Append(indentation6);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold13 = new Bold();

            runProperties8.Append(runFonts14);
            runProperties8.Append(bold13);
            Text text8 = new Text();
            text8.Text = "--------------------------------------";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold14 = new Bold();

            runProperties9.Append(runFonts15);
            runProperties9.Append(bold14);
            Text text9 = new Text();
            text9.Text = "---";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run8);
            paragraph6.Append(run9);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);
            tableCell4.Append(paragraph6);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "004632DA", RsidParagraphProperties = "004632DA", RsidRunAdditionDefault = "004632DA", ParagraphId = "056DFB6D", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation7 = new Indentation() { End = "-357" };
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold15 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties7.Append(runFonts16);
            paragraphMarkRunProperties7.Append(bold15);
            paragraphMarkRunProperties7.Append(fontSize4);

            paragraphProperties7.Append(spacingBetweenLines7);
            paragraphProperties7.Append(indentation7);
            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            paragraph7.Append(paragraphProperties7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "004632DA", RsidParagraphAddition = "00043D04", RsidParagraphProperties = "004632DA", RsidRunAdditionDefault = "004632DA", ParagraphId = "19EDA5C0", TextId = "2198FEA6" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation8 = new Indentation() { End = "-357" };
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold16 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties8.Append(runFonts17);
            paragraphMarkRunProperties8.Append(bold16);
            paragraphMarkRunProperties8.Append(fontSize5);

            paragraphProperties8.Append(spacingBetweenLines8);
            paragraphProperties8.Append(indentation8);
            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run10 = new Run() { RsidRunProperties = "004632DA" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold17 = new Bold();
            FontSize fontSize6 = new FontSize() { Val = "28" };

            runProperties10.Append(runFonts18);
            runProperties10.Append(bold17);
            runProperties10.Append(fontSize6);
            Text text10 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text10.Text = "THÔNG TIN GIẢNG DẠY LÝ THUYẾT, BÀI TẬP VÀ THỰC HÀNH ";

            run10.Append(runProperties10);
            run10.Append(text10);

            Run run11 = new Run() { RsidRunProperties = "004632DA" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold18 = new Bold();
            FontSize fontSize7 = new FontSize() { Val = "28" };

            runProperties11.Append(runFonts19);
            runProperties11.Append(bold18);
            runProperties11.Append(fontSize7);
            Break break1 = new Break();
            Text text11 = new Text();
            text11.Text = "TẠI TRƯỜNG ĐẠI HỌC KHOA HỌC TỰ NHIÊN";

            run11.Append(runProperties11);
            run11.Append(break1);
            run11.Append(text11);

            Run run12 = new Run() { RsidRunProperties = "004632DA", RsidRunAddition = "00043D04" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold19 = new Bold();
            FontSize fontSize8 = new FontSize() { Val = "28" };

            runProperties12.Append(runFonts20);
            runProperties12.Append(bold19);
            runProperties12.Append(fontSize8);
            Break break2 = new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.All };
            Text text12 = new Text();
            text12.Text = "Năm học 2020-2021";

            run12.Append(runProperties12);
            run12.Append(break2);
            run12.Append(text12);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run10);
            paragraph8.Append(run11);
            paragraph8.Append(run12);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "004632DA", RsidParagraphProperties = "00043D04", RsidRunAdditionDefault = "004632DA", ParagraphId = "0094EBD1", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs1.Append(tabStop1);
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation9 = new Indentation() { Start = "720", End = "-357" };
            Justification justification9 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize9 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties9.Append(runFonts21);
            paragraphMarkRunProperties9.Append(boldComplexScript1);
            paragraphMarkRunProperties9.Append(italicComplexScript1);
            paragraphMarkRunProperties9.Append(fontSize9);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript4);

            paragraphProperties9.Append(tabs1);
            paragraphProperties9.Append(spacingBetweenLines9);
            paragraphProperties9.Append(indentation9);
            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            paragraph9.Append(paragraphProperties9);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "004632DA", RsidParagraphAddition = "00043D04", RsidParagraphProperties = "00043D04", RsidRunAdditionDefault = "00043D04", ParagraphId = "06A9CC42", TextId = "67534542" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs2.Append(tabStop2);
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation10 = new Indentation() { Start = "720", End = "-357" };
            Justification justification10 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold20 = new Bold();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            FontSize fontSize10 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties10.Append(runFonts22);
            paragraphMarkRunProperties10.Append(bold20);
            paragraphMarkRunProperties10.Append(italicComplexScript2);
            paragraphMarkRunProperties10.Append(fontSize10);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript5);

            paragraphProperties10.Append(tabs2);
            paragraphProperties10.Append(spacingBetweenLines10);
            paragraphProperties10.Append(indentation10);
            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run13 = new Run() { RsidRunProperties = "004632DA" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold21 = new Bold();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
            FontSize fontSize11 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "26" };

            runProperties13.Append(runFonts23);
            runProperties13.Append(bold21);
            runProperties13.Append(italicComplexScript3);
            runProperties13.Append(fontSize11);
            runProperties13.Append(fontSizeComplexScript6);
            Text text13 = new Text();
            text13.Text = "Quý Thầy Cô: \t" + gv.Ten;

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run13);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "004632DA", RsidParagraphAddition = "00043D04", RsidParagraphProperties = "00043D04", RsidRunAdditionDefault = "00043D04", ParagraphId = "45EC630A", TextId = "1BF9C010" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs3.Append(tabStop3);
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation11 = new Indentation() { Start = "720", End = "-357" };
            Justification justification11 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold22 = new Bold();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            FontSize fontSize12 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties11.Append(runFonts24);
            paragraphMarkRunProperties11.Append(bold22);
            paragraphMarkRunProperties11.Append(italicComplexScript4);
            paragraphMarkRunProperties11.Append(fontSize12);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript7);

            paragraphProperties11.Append(tabs3);
            paragraphProperties11.Append(spacingBetweenLines11);
            paragraphProperties11.Append(indentation11);
            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run14 = new Run() { RsidRunProperties = "004632DA" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold23 = new Bold();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            FontSize fontSize13 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "26" };

            runProperties14.Append(runFonts25);
            runProperties14.Append(bold23);
            runProperties14.Append(italicComplexScript5);
            runProperties14.Append(fontSize13);
            runProperties14.Append(fontSizeComplexScript8);
            Text text14 = new Text();
            text14.Text = "Mã số cán bộ: \t" + gv.MSCB;

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run14);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00043D04", RsidParagraphAddition = "00043D04", RsidParagraphProperties = "00043D04", RsidRunAdditionDefault = "00043D04", ParagraphId = "5749ECD8", TextId = "77777777" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs4.Append(tabStop4);
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation12 = new Indentation() { Start = "720", End = "-357" };
            Justification justification12 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript6 = new ItalicComplexScript();
            FontSize fontSize14 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties12.Append(runFonts26);
            paragraphMarkRunProperties12.Append(boldComplexScript2);
            paragraphMarkRunProperties12.Append(italicComplexScript6);
            paragraphMarkRunProperties12.Append(fontSize14);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript9);

            paragraphProperties12.Append(tabs4);
            paragraphProperties12.Append(spacingBetweenLines12);
            paragraphProperties12.Append(indentation12);
            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            paragraph12.Append(paragraphProperties12);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "000F1E75", RsidRunAdditionDefault = "000F1E75", ParagraphId = "7690B51C", TextId = "7655200A" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs5.Append(tabStop5);
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation13 = new Indentation() { End = "-357" };
            Justification justification13 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold24 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript7 = new ItalicComplexScript();
            FontSize fontSize15 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties13.Append(runFonts27);
            paragraphMarkRunProperties13.Append(bold24);
            paragraphMarkRunProperties13.Append(boldComplexScript3);
            paragraphMarkRunProperties13.Append(italicComplexScript7);
            paragraphMarkRunProperties13.Append(fontSize15);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript10);

            paragraphProperties13.Append(tabs5);
            paragraphProperties13.Append(spacingBetweenLines13);
            paragraphProperties13.Append(indentation13);
            paragraphProperties13.Append(justification13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run15 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript8 = new ItalicComplexScript();
            FontSize fontSize16 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "26" };

            runProperties15.Append(runFonts28);
            runProperties15.Append(boldComplexScript4);
            runProperties15.Append(italicComplexScript8);
            runProperties15.Append(fontSize16);
            runProperties15.Append(fontSizeComplexScript11);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = "Dưới đây là ";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript9 = new ItalicComplexScript();
            FontSize fontSize17 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "26" };

            runProperties16.Append(runFonts29);
            runProperties16.Append(boldComplexScript5);
            runProperties16.Append(italicComplexScript9);
            runProperties16.Append(fontSize17);
            runProperties16.Append(fontSizeComplexScript12);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = "thông tin giảng dạy Đại học và Sau đại học tại Trường Đại học Khoa học Tự nhiên ";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript10 = new ItalicComplexScript();
            FontSize fontSize18 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "26" };

            runProperties17.Append(runFonts30);
            runProperties17.Append(boldComplexScript6);
            runProperties17.Append(italicComplexScript10);
            runProperties17.Append(fontSize18);
            runProperties17.Append(fontSizeComplexScript13);
            Text text17 = new Text();
            text17.Text = "mà hệ thống nhà trường đã ghi nhận.";

            run17.Append(runProperties17);
            run17.Append(text17);

            Run run18 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript11 = new ItalicComplexScript();
            FontSize fontSize19 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "26" };

            runProperties18.Append(runFonts31);
            runProperties18.Append(boldComplexScript7);
            runProperties18.Append(italicComplexScript11);
            runProperties18.Append(fontSize19);
            runProperties18.Append(fontSizeComplexScript14);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = " Quý Thầy Cô sử dụng thông tin này để đưa vào báo cáo của mình";

            run18.Append(runProperties18);
            run18.Append(text18);

            Run run19 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript12 = new ItalicComplexScript();
            FontSize fontSize20 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "26" };

            runProperties19.Append(runFonts32);
            runProperties19.Append(boldComplexScript8);
            runProperties19.Append(italicComplexScript12);
            runProperties19.Append(fontSize20);
            runProperties19.Append(fontSizeComplexScript15);
            Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text19.Text = " trong mục ";

            run19.Append(runProperties19);
            run19.Append(text19);

            Run run20 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold25 = new Bold();
            ItalicComplexScript italicComplexScript13 = new ItalicComplexScript();
            FontSize fontSize21 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "26" };

            runProperties20.Append(runFonts33);
            runProperties20.Append(bold25);
            runProperties20.Append(italicComplexScript13);
            runProperties20.Append(fontSize21);
            runProperties20.Append(fontSizeComplexScript16);
            Text text20 = new Text();
            text20.Text = "(a)";

            run20.Append(runProperties20);
            run20.Append(text20);

            Run run21 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript14 = new ItalicComplexScript();
            FontSize fontSize22 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "26" };

            runProperties21.Append(runFonts34);
            runProperties21.Append(boldComplexScript9);
            runProperties21.Append(italicComplexScript14);
            runProperties21.Append(fontSize22);
            runProperties21.Append(fontSizeComplexScript17);
            Text text21 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text21.Text = " ";

            run21.Append(runProperties21);
            run21.Append(text21);

            Run run22 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold26 = new Bold();
            ItalicComplexScript italicComplexScript15 = new ItalicComplexScript();
            FontSize fontSize23 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "26" };

            runProperties22.Append(runFonts35);
            runProperties22.Append(bold26);
            runProperties22.Append(italicComplexScript15);
            runProperties22.Append(fontSize23);
            runProperties22.Append(fontSizeComplexScript18);
            Text text22 = new Text();
            text22.Text = "Thông tin giảng dạy";

            run22.Append(runProperties22);
            run22.Append(text22);

            Run run23 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript16 = new ItalicComplexScript();
            FontSize fontSize24 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "26" };

            runProperties23.Append(runFonts36);
            runProperties23.Append(boldComplexScript10);
            runProperties23.Append(italicComplexScript16);
            runProperties23.Append(fontSize24);
            runProperties23.Append(fontSizeComplexScript19);
            Text text23 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text23.Text = " ";

            run23.Append(runProperties23);
            run23.Append(text23);

            Run run24 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold27 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript17 = new ItalicComplexScript();
            FontSize fontSize25 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "26" };

            runProperties24.Append(runFonts37);
            runProperties24.Append(bold27);
            runProperties24.Append(boldComplexScript11);
            runProperties24.Append(italicComplexScript17);
            runProperties24.Append(fontSize25);
            runProperties24.Append(fontSizeComplexScript20);
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = "tại Trường Đại học Khoa học Tự nhiên, ";

            run24.Append(runProperties24);
            run24.Append(text24);

            Run run25 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            ItalicComplexScript italicComplexScript18 = new ItalicComplexScript();
            FontSize fontSize26 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "26" };

            runProperties25.Append(runFonts38);
            runProperties25.Append(italicComplexScript18);
            runProperties25.Append(fontSize26);
            runProperties25.Append(fontSizeComplexScript21);
            Text text25 = new Text();
            text25.Text = "trong phần";

            run25.Append(runProperties25);
            run25.Append(text25);

            Run run26 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold28 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript19 = new ItalicComplexScript();
            FontSize fontSize27 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "26" };

            runProperties26.Append(runFonts39);
            runProperties26.Append(bold28);
            runProperties26.Append(boldComplexScript12);
            runProperties26.Append(italicComplexScript19);
            runProperties26.Append(fontSize27);
            runProperties26.Append(fontSizeComplexScript22);
            Text text26 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text26.Text = " Giảng dạy lý thuyết, bài tập và thực hành";

            run26.Append(runProperties26);
            run26.Append(text26);

            Run run27 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold29 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript20 = new ItalicComplexScript();
            FontSize fontSize28 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "26" };

            runProperties27.Append(runFonts40);
            runProperties27.Append(bold29);
            runProperties27.Append(boldComplexScript13);
            runProperties27.Append(italicComplexScript20);
            runProperties27.Append(fontSize28);
            runProperties27.Append(fontSizeComplexScript23);
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = " ";

            run27.Append(runProperties27);
            run27.Append(text27);

            Run run28 = new Run() { RsidRunProperties = "00485A86", RsidRunAddition = "00485A86" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            ItalicComplexScript italicComplexScript21 = new ItalicComplexScript();
            FontSize fontSize29 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "26" };

            runProperties28.Append(runFonts41);
            runProperties28.Append(italicComplexScript21);
            runProperties28.Append(fontSize29);
            runProperties28.Append(fontSizeComplexScript24);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = "của ";

            run28.Append(runProperties28);
            run28.Append(text28);

            Run run29 = new Run() { RsidRunProperties = "00485A86", RsidRunAddition = "00485A86" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold30 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript22 = new ItalicComplexScript();
            FontSize fontSize30 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "26" };

            runProperties29.Append(runFonts42);
            runProperties29.Append(bold30);
            runProperties29.Append(boldComplexScript14);
            runProperties29.Append(italicComplexScript22);
            runProperties29.Append(fontSize30);
            runProperties29.Append(fontSizeComplexScript25);
            Text text29 = new Text();
            text29.Text = "Báo cáo công tác năm học";

            run29.Append(runProperties29);
            run29.Append(text29);

            Run run30 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold31 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript23 = new ItalicComplexScript();
            FontSize fontSize31 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "26" };

            runProperties30.Append(runFonts43);
            runProperties30.Append(bold31);
            runProperties30.Append(boldComplexScript15);
            runProperties30.Append(italicComplexScript23);
            runProperties30.Append(fontSize31);
            runProperties30.Append(fontSizeComplexScript26);
            Text text30 = new Text();
            text30.Text = ".";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run15);
            paragraph13.Append(run16);
            paragraph13.Append(run17);
            paragraph13.Append(run18);
            paragraph13.Append(run19);
            paragraph13.Append(run20);
            paragraph13.Append(run21);
            paragraph13.Append(run22);
            paragraph13.Append(run23);
            paragraph13.Append(run24);
            paragraph13.Append(run25);
            paragraph13.Append(run26);
            paragraph13.Append(run27);
            paragraph13.Append(run28);
            paragraph13.Append(run29);
            paragraph13.Append(run30);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "000F1E75", RsidRunAdditionDefault = "000F1E75", ParagraphId = "7307B4F5", TextId = "77777777" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();

            Tabs tabs6 = new Tabs();
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs6.Append(tabStop6);
            Indentation indentation14 = new Indentation() { End = "-357" };
            Justification justification14 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            Italic italic1 = new Italic();
            FontSize fontSize32 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties14.Append(runFonts44);
            paragraphMarkRunProperties14.Append(boldComplexScript16);
            paragraphMarkRunProperties14.Append(italic1);
            paragraphMarkRunProperties14.Append(fontSize32);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript27);

            paragraphProperties14.Append(tabs6);
            paragraphProperties14.Append(indentation14);
            paragraphProperties14.Append(justification14);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            paragraph14.Append(paragraphProperties14);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "9730", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification4 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };
            TableLook tableLook2 = new TableLook() { Val = "0000" };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableJustification4);
            tableProperties2.Append(tableBorders1);
            tableProperties2.Append(tableLayout1);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn3 = new GridColumn() { Width = "586" };
            GridColumn gridColumn4 = new GridColumn() { Width = "3348" };
            GridColumn gridColumn5 = new GridColumn() { Width = "1000" };
            GridColumn gridColumn6 = new GridColumn() { Width = "3350" };
            GridColumn gridColumn7 = new GridColumn() { Width = "1446" };

            tableGrid2.Append(gridColumn3);
            tableGrid2.Append(gridColumn4);
            tableGrid2.Append(gridColumn5);
            tableGrid2.Append(gridColumn6);
            tableGrid2.Append(gridColumn7);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00B42C64", RsidTableRowAddition = "000F1E75", RsidTableRowProperties = "00FF498E", ParagraphId = "2FFA4C76", TextId = "77777777" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableJustification tableJustification5 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties3.Append(tableJustification5);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "586", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders1);
            tableCellProperties5.Append(tableCellVerticalAlignment1);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "09B60016", TextId = "77777777" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading3" };
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation15 = new Indentation() { End = "0" };
            Justification justification15 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic2 = new Italic() { Val = false };
            FontSize fontSize33 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties15.Append(runFonts45);
            paragraphMarkRunProperties15.Append(italic2);
            paragraphMarkRunProperties15.Append(fontSize33);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript28);

            paragraphProperties15.Append(paragraphStyleId1);
            paragraphProperties15.Append(spacingBetweenLines14);
            paragraphProperties15.Append(indentation15);
            paragraphProperties15.Append(justification15);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run31 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic3 = new Italic() { Val = false };
            FontSize fontSize34 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "26" };

            runProperties31.Append(runFonts46);
            runProperties31.Append(italic3);
            runProperties31.Append(fontSize34);
            runProperties31.Append(fontSizeComplexScript29);
            Text text31 = new Text();
            text31.Text = "TT";

            run31.Append(runProperties31);
            run31.Append(text31);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run31);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph15);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "3348", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);
            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(rightBorder3);
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders2);
            tableCellProperties6.Append(tableCellVerticalAlignment2);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "262E00D2", TextId = "77777777" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Heading7" };
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation16 = new Indentation() { End = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic4 = new Italic() { Val = false };
            FontSize fontSize35 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties16.Append(runFonts47);
            paragraphMarkRunProperties16.Append(italic4);
            paragraphMarkRunProperties16.Append(fontSize35);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript30);

            paragraphProperties16.Append(paragraphStyleId2);
            paragraphProperties16.Append(spacingBetweenLines15);
            paragraphProperties16.Append(indentation16);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run32 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic5 = new Italic() { Val = false };
            FontSize fontSize36 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "26" };

            runProperties32.Append(runFonts48);
            runProperties32.Append(italic5);
            runProperties32.Append(fontSize36);
            runProperties32.Append(fontSizeComplexScript31);
            Text text32 = new Text();
            text32.Text = "Môn học";

            run32.Append(runProperties32);
            run32.Append(text32);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run32);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph16);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(topBorder4);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder4);
            tableCellBorders3.Append(rightBorder4);
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders3);
            tableCellProperties7.Append(tableCellVerticalAlignment3);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "4DE22DF0", TextId = "77777777" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Heading6" };
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation17 = new Indentation() { End = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic6 = new Italic() { Val = false };
            FontSize fontSize37 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties17.Append(runFonts49);
            paragraphMarkRunProperties17.Append(italic6);
            paragraphMarkRunProperties17.Append(fontSize37);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript32);

            paragraphProperties17.Append(paragraphStyleId3);
            paragraphProperties17.Append(spacingBetweenLines16);
            paragraphProperties17.Append(indentation17);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run33 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic7 = new Italic() { Val = false };
            FontSize fontSize38 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "26" };

            runProperties33.Append(runFonts50);
            runProperties33.Append(italic7);
            runProperties33.Append(fontSize38);
            runProperties33.Append(fontSizeComplexScript33);
            Text text33 = new Text();
            text33.Text = "Lớp";

            run33.Append(runProperties33);
            run33.Append(text33);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run33);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph17);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "3350", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder5);
            tableCellBorders4.Append(leftBorder5);
            tableCellBorders4.Append(bottomBorder5);
            tableCellBorders4.Append(rightBorder5);
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellBorders4);
            tableCellProperties8.Append(tableCellVerticalAlignment4);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "79092C1F", TextId = "77777777" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "Heading4" };
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation18 = new Indentation() { Start = "0", End = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic8 = new Italic() { Val = false };
            FontSize fontSize39 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties18.Append(runFonts51);
            paragraphMarkRunProperties18.Append(italic8);
            paragraphMarkRunProperties18.Append(fontSize39);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript34);

            paragraphProperties18.Append(paragraphStyleId4);
            paragraphProperties18.Append(spacingBetweenLines17);
            paragraphProperties18.Append(indentation18);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run34 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic9 = new Italic() { Val = false };
            FontSize fontSize40 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "26" };

            runProperties34.Append(runFonts52);
            runProperties34.Append(italic9);
            runProperties34.Append(fontSize40);
            runProperties34.Append(fontSizeComplexScript35);
            Text text34 = new Text();
            text34.Text = "Bậc đào tạo";

            run34.Append(runProperties34);
            run34.Append(text34);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run34);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph18);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1446", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder6);
            tableCellBorders5.Append(leftBorder6);
            tableCellBorders5.Append(bottomBorder6);
            tableCellBorders5.Append(rightBorder6);
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellBorders5);
            tableCellProperties9.Append(tableCellVerticalAlignment5);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "6FEE9C97", TextId = "77777777" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "Heading4" };
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation19 = new Indentation() { Start = "0", End = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold32 = new Bold() { Val = false };
            FontSize fontSize41 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties19.Append(runFonts53);
            paragraphMarkRunProperties19.Append(bold32);
            paragraphMarkRunProperties19.Append(fontSize41);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript36);

            paragraphProperties19.Append(paragraphStyleId5);
            paragraphProperties19.Append(spacingBetweenLines18);
            paragraphProperties19.Append(indentation19);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run35 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Italic italic10 = new Italic() { Val = false };
            FontSize fontSize42 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "26" };

            runProperties35.Append(runFonts54);
            runProperties35.Append(italic10);
            runProperties35.Append(fontSize42);
            runProperties35.Append(fontSizeComplexScript37);
            Text text35 = new Text();
            text35.Text = "Giờ chuẩn";

            run35.Append(runProperties35);
            run35.Append(text35);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run35);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph19);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);

            // Nghia
            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow3);
            // CODE HERE - INSERTS ROWS TO TABLE
            double cong = 0;
            for (int i = 0; i < gv.Ttgd.Count; i++)
            {
                ThongTinGiangDay currentTT = gv.Ttgd[i];
                cong += currentTT.GioChuan;
                TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00B42C64", RsidTableRowAddition = "000F1E75", RsidTableRowProperties = "00FF498E", ParagraphId = "266C0062", TextId = "77777777" };

                TableRowProperties tableRowProperties4 = new TableRowProperties();
                TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)372U };
                TableJustification tableJustification6 = new TableJustification() { Val = TableRowAlignmentValues.Center };

                tableRowProperties4.Append(tableRowHeight1);
                tableRowProperties4.Append(tableJustification6);

                TableCell tableCell10 = new TableCell();

                TableCellProperties tableCellProperties10 = new TableCellProperties();
                TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "586", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders6 = new TableCellBorders();
                TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders6.Append(topBorder7);
                tableCellBorders6.Append(leftBorder7);
                tableCellBorders6.Append(bottomBorder7);
                tableCellBorders6.Append(rightBorder7);
                TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties10.Append(tableCellWidth10);
                tableCellProperties10.Append(tableCellBorders6);
                tableCellProperties10.Append(tableCellVerticalAlignment6);

                Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00474011", ParagraphId = "30540B3B", TextId = "30228281" };

                ParagraphProperties paragraphProperties20 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
                Justification justification16 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
                RunFonts nrunFonts50 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize nfontSize39 = new FontSize() { Val = "26" };
                FontSizeComplexScript nfontSizeComplexScript34 = new FontSizeComplexScript() { Val = "26" };

                paragraphMarkRunProperties20.Append(nrunFonts50);
                paragraphMarkRunProperties20.Append(nfontSize39);
                paragraphMarkRunProperties20.Append(nfontSizeComplexScript34);

                paragraphProperties20.Append(spacingBetweenLines19);
                paragraphProperties20.Append(justification16);
                paragraphProperties20.Append(paragraphMarkRunProperties20);

                Run nrun31 = new Run();

                RunProperties nrunProperties31 = new RunProperties();
                RunFonts nrunFonts51 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize nfontSize40 = new FontSize() { Val = "26" };
                FontSizeComplexScript nfontSizeComplexScript35 = new FontSizeComplexScript() { Val = "26" };

                nrunProperties31.Append(nrunFonts51);
                nrunProperties31.Append(nfontSize40);
                nrunProperties31.Append(nfontSizeComplexScript35);
                Text ntext31 = new Text();
                ntext31.Text = "";

                nrun31.Append(nrunProperties31);
                nrun31.Append(ntext31);

                Run nrun32 = new Run() { RsidRunProperties = "00B42C64", RsidRunAddition = "000F1E75" };

                RunProperties nrunProperties32 = new RunProperties();
                RunFonts nrunFonts52 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize nfontSize41 = new FontSize() { Val = "26" };
                FontSizeComplexScript nfontSizeComplexScript36 = new FontSizeComplexScript() { Val = "26" };

                runProperties32.Append(nrunFonts52);
                runProperties32.Append(nfontSize41);
                runProperties32.Append(nfontSizeComplexScript36);
                Text ntext32 = new Text();
                ntext32.Text = (i + 1).ToString();

                nrun32.Append(nrunProperties32);
                nrun32.Append(ntext32);

                paragraph20.Append(paragraphProperties20);
                paragraph20.Append(nrun31);
                paragraph20.Append(nrun32);

                tableCell10.Append(tableCellProperties10);
                tableCell10.Append(paragraph20);

                TableCell tableCell11 = new TableCell();

                TableCellProperties tableCellProperties11 = new TableCellProperties();
                TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "3348", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders7 = new TableCellBorders();
                TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders7.Append(topBorder8);
                tableCellBorders7.Append(leftBorder8);
                tableCellBorders7.Append(bottomBorder8);
                tableCellBorders7.Append(rightBorder8);

                tableCellProperties11.Append(tableCellWidth11);
                tableCellProperties11.Append(tableCellBorders7);

                Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00362B77", ParagraphId = "233FB166", TextId = "36B1435A" };

                ParagraphProperties paragraphProperties21 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
                RunFonts nrunFonts53 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize nfontSize42 = new FontSize() { Val = "26" };
                FontSizeComplexScript nfontSizeComplexScript37 = new FontSizeComplexScript() { Val = "26" };

                paragraphMarkRunProperties21.Append(nrunFonts53);
                paragraphMarkRunProperties21.Append(nfontSize42);
                paragraphMarkRunProperties21.Append(nfontSizeComplexScript37);

                paragraphProperties21.Append(spacingBetweenLines20);
                paragraphProperties21.Append(paragraphMarkRunProperties21);

                Run nrun33 = new Run();

                RunProperties nrunProperties33 = new RunProperties();
                RunFonts nrunFonts54 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize43 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "26" };

                nrunProperties33.Append(nrunFonts54);
                nrunProperties33.Append(fontSize43);
                nrunProperties33.Append(fontSizeComplexScript38);
                Text ntext33 = new Text();
                ntext33.Text = currentTT.MonHoc;

                nrun33.Append(nrunProperties33);
                nrun33.Append(ntext33);

                paragraph21.Append(paragraphProperties21);
                paragraph21.Append(nrun33);

                tableCell11.Append(tableCellProperties11);
                tableCell11.Append(paragraph21);

                TableCell tableCell12 = new TableCell();

                TableCellProperties tableCellProperties12 = new TableCellProperties();
                TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1000", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders8 = new TableCellBorders();
                TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders8.Append(topBorder9);
                tableCellBorders8.Append(leftBorder9);
                tableCellBorders8.Append(bottomBorder9);
                tableCellBorders8.Append(rightBorder9);

                tableCellProperties12.Append(tableCellWidth12);
                tableCellProperties12.Append(tableCellBorders8);

                Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00362B77", ParagraphId = "5D855D29", TextId = "03665D11" };

                ParagraphProperties paragraphProperties22 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
                RunFonts runFonts55 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize44 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "26" };

                paragraphMarkRunProperties22.Append(runFonts55);
                paragraphMarkRunProperties22.Append(fontSize44);
                paragraphMarkRunProperties22.Append(fontSizeComplexScript39);

                paragraphProperties22.Append(spacingBetweenLines21);
                paragraphProperties22.Append(paragraphMarkRunProperties22);

                Run nrun34 = new Run();

                RunProperties nrunProperties34 = new RunProperties();
                RunFonts runFonts56 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize45 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "26" };

                nrunProperties34.Append(runFonts56);
                nrunProperties34.Append(fontSize45);
                nrunProperties34.Append(fontSizeComplexScript40);
                Text ntext34 = new Text();
                ntext34.Text = currentTT.Lop;

                nrun34.Append(nrunProperties34);
                nrun34.Append(ntext34);

                paragraph22.Append(paragraphProperties22);
                paragraph22.Append(nrun34);

                tableCell12.Append(tableCellProperties12);
                tableCell12.Append(paragraph22);

                TableCell tableCell13 = new TableCell();

                TableCellProperties tableCellProperties13 = new TableCellProperties();
                TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "3350", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders9 = new TableCellBorders();
                TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders9.Append(topBorder10);
                tableCellBorders9.Append(leftBorder10);
                tableCellBorders9.Append(bottomBorder10);
                tableCellBorders9.Append(rightBorder10);

                tableCellProperties13.Append(tableCellWidth13);
                tableCellProperties13.Append(tableCellBorders9);

                Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00362B77", ParagraphId = "6053D631", TextId = "60876A37" };

                ParagraphProperties paragraphProperties23 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
                RunFonts runFonts57 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize46 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "26" };

                paragraphMarkRunProperties23.Append(runFonts57);
                paragraphMarkRunProperties23.Append(fontSize46);
                paragraphMarkRunProperties23.Append(fontSizeComplexScript41);

                paragraphProperties23.Append(spacingBetweenLines22);
                paragraphProperties23.Append(paragraphMarkRunProperties23);

                Run nrun35 = new Run();

                RunProperties nrunProperties35 = new RunProperties();
                RunFonts runFonts58 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize47 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "26" };

                nrunProperties35.Append(runFonts58);
                nrunProperties35.Append(fontSize47);
                nrunProperties35.Append(fontSizeComplexScript42);
                Text ntext35 = new Text();
                ntext35.Text = currentTT.BacDaoTao;

                nrun35.Append(nrunProperties35);
                nrun35.Append(ntext35);

                paragraph23.Append(paragraphProperties23);
                paragraph23.Append(nrun35);

                tableCell13.Append(tableCellProperties13);
                tableCell13.Append(paragraph23);

                TableCell tableCell14 = new TableCell();

                TableCellProperties tableCellProperties14 = new TableCellProperties();
                TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1446", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders10 = new TableCellBorders();
                TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
                RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

                tableCellBorders10.Append(topBorder11);
                tableCellBorders10.Append(leftBorder11);
                tableCellBorders10.Append(bottomBorder11);
                tableCellBorders10.Append(rightBorder11);

                tableCellProperties14.Append(tableCellWidth14);
                tableCellProperties14.Append(tableCellBorders10);

                Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "00362B77", ParagraphId = "0B4D1003", TextId = "0F161137" };

                ParagraphProperties paragraphProperties24 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
                RunFonts runFonts59 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize48 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "26" };

                paragraphMarkRunProperties24.Append(runFonts59);
                paragraphMarkRunProperties24.Append(fontSize48);
                paragraphMarkRunProperties24.Append(fontSizeComplexScript43);

                paragraphProperties24.Append(spacingBetweenLines23);
                paragraphProperties24.Append(paragraphMarkRunProperties24);

                Run run36 = new Run();

                RunProperties runProperties36 = new RunProperties();
                RunFonts runFonts60 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize49 = new FontSize() { Val = "26" };
                FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "26" };

                runProperties36.Append(runFonts60);
                runProperties36.Append(fontSize49);
                runProperties36.Append(fontSizeComplexScript44);
                Text text36 = new Text();
                text36.Text = currentTT.GioChuan.ToString();

                run36.Append(runProperties36);
                run36.Append(text36);

                paragraph24.Append(paragraphProperties24);
                paragraph24.Append(run36);

                tableCell14.Append(tableCellProperties14);
                tableCell14.Append(paragraph24);

                tableRow4.Append(tableRowProperties4);
                tableRow4.Append(tableCell10);
                tableRow4.Append(tableCell11);
                tableRow4.Append(tableCell12);
                tableRow4.Append(tableCell13);
                tableRow4.Append(tableCell14);

                //add row
                table2.Append(tableRow4);
            }

            // Table row sum
            TableRow tableRow7 = new TableRow() { RsidTableRowMarkRevision = "00B42C64", RsidTableRowAddition = "000F1E75", RsidTableRowProperties = "00FF498E", ParagraphId = "7D4D1DC7", TextId = "77777777" };

            TableRowProperties tableRowProperties7 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)296U };
            TableJustification tableJustification9 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties7.Append(tableRowHeight4);
            tableRowProperties7.Append(tableJustification9);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "586", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders21 = new TableCellBorders();
            TopBorder topBorder22 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder22 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders21.Append(topBorder22);
            tableCellBorders21.Append(leftBorder22);
            tableCellBorders21.Append(bottomBorder22);
            tableCellBorders21.Append(rightBorder22);

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(tableCellBorders21);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "6125A975", TextId = "77777777" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines34 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize68 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties35.Append(runFonts79);
            paragraphMarkRunProperties35.Append(fontSize68);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript63);

            paragraphProperties35.Append(spacingBetweenLines34);
            paragraphProperties35.Append(justification19);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            paragraph35.Append(paragraphProperties35);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph35);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "3348", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders22 = new TableCellBorders();
            TopBorder topBorder23 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder23 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder23 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder23 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders22.Append(topBorder23);
            tableCellBorders22.Append(leftBorder23);
            tableCellBorders22.Append(bottomBorder23);
            tableCellBorders22.Append(rightBorder23);

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(tableCellBorders22);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "7CE2E6A9", TextId = "77777777" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines35 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold34 = new Bold();
            FontSize fontSize69 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties36.Append(runFonts80);
            paragraphMarkRunProperties36.Append(bold34);
            paragraphMarkRunProperties36.Append(fontSize69);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript64);

            paragraphProperties36.Append(spacingBetweenLines35);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run45 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold35 = new Bold();
            FontSize fontSize70 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "26" };

            runProperties45.Append(runFonts81);
            runProperties45.Append(bold35);
            runProperties45.Append(fontSize70);
            runProperties45.Append(fontSizeComplexScript65);
            Text text45 = new Text();
            text45.Text = "Cộng";

            run45.Append(runProperties45);
            run45.Append(text45);

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run45);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph36);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "1000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders23 = new TableCellBorders();
            TopBorder topBorder24 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder24 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder24 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder24 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders23.Append(topBorder24);
            tableCellBorders23.Append(leftBorder24);
            tableCellBorders23.Append(bottomBorder24);
            tableCellBorders23.Append(rightBorder24);

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(tableCellBorders23);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "27B7D696", TextId = "77777777" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines36 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize71 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties37.Append(runFonts82);
            paragraphMarkRunProperties37.Append(fontSize71);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript66);

            paragraphProperties37.Append(spacingBetweenLines36);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            paragraph37.Append(paragraphProperties37);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph37);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "3350", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders24 = new TableCellBorders();
            TopBorder topBorder25 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder25 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder25 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder25 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders24.Append(topBorder25);
            tableCellBorders24.Append(leftBorder25);
            tableCellBorders24.Append(bottomBorder25);
            tableCellBorders24.Append(rightBorder25);

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(tableCellBorders24);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "52AFEF44", TextId = "77777777" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines37 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts83 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize72 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties38.Append(runFonts83);
            paragraphMarkRunProperties38.Append(fontSize72);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript67);

            paragraphProperties38.Append(spacingBetweenLines37);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph38);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "1446", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders25 = new TableCellBorders();
            TopBorder topBorder26 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder26 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder26 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder26 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders25.Append(topBorder26);
            tableCellBorders25.Append(leftBorder26);
            tableCellBorders25.Append(bottomBorder26);
            tableCellBorders25.Append(rightBorder26);

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(tableCellBorders25);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "00B42C64", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00FF498E", RsidRunAdditionDefault = "000F1E75", ParagraphId = "427AD513", TextId = "77777777" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines38 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts84 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize73 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties39.Append(runFonts84);
            paragraphMarkRunProperties39.Append(fontSize73);
            paragraphMarkRunProperties39.Append(fontSizeComplexScript68);

            paragraphProperties39.Append(spacingBetweenLines38);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run46 = new Run() { RsidRunProperties = "00B42C64" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize74 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "26" };

            runProperties46.Append(runFonts85);
            runProperties46.Append(fontSize74);
            runProperties46.Append(fontSizeComplexScript69);
            Text text46 = new Text();
            text46.Text = cong.ToString();

            run46.Append(runProperties46);
            run46.Append(text46);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run46);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph39);

            tableRow7.Append(tableRowProperties7);
            tableRow7.Append(tableCell25);
            tableRow7.Append(tableCell26);
            tableRow7.Append(tableCell27);
            tableRow7.Append(tableCell28);
            tableRow7.Append(tableCell29);

            table2.Append(tableRow7);



            // END CODE HERE


            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "000F1E75", RsidRunAdditionDefault = "000F1E75", ParagraphId = "127C605E", TextId = "59F41413" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs7.Append(tabStop7);
            SpacingBetweenLines spacingBetweenLines39 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation20 = new Indentation() { End = "-357" };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts nrunFonts83 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold37 = new Bold();
            Italic italic11 = new Italic();
            FontSize nfontSize71 = new FontSize() { Val = "26" };
            FontSizeComplexScript nfontSizeComplexScript66 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties40.Append(nrunFonts83);
            paragraphMarkRunProperties40.Append(bold37);
            paragraphMarkRunProperties40.Append(italic11);
            paragraphMarkRunProperties40.Append(nfontSize71);
            paragraphMarkRunProperties40.Append(nfontSizeComplexScript66);

            paragraphProperties40.Append(tabs7);
            paragraphProperties40.Append(spacingBetweenLines39);
            paragraphProperties40.Append(indentation20);
            paragraphProperties40.Append(paragraphMarkRunProperties40);

            paragraph40.Append(paragraphProperties40);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphAddition = "00485A86", RsidParagraphProperties = "00485A86", RsidRunAdditionDefault = "000F1E75", ParagraphId = "558223CD", TextId = "081211F2" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();

            Tabs tabs8 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs8.Append(tabStop8);
            SpacingBetweenLines spacingBetweenLines40 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation21 = new Indentation() { End = "-357" };
            Justification justification20 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts nrunFonts84 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold38 = new Bold();
            BoldComplexScript boldComplexScript17 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript24 = new ItalicComplexScript();
            FontSize nfontSize72 = new FontSize() { Val = "26" };
            FontSizeComplexScript nfontSizeComplexScript67 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties41.Append(nrunFonts84);
            paragraphMarkRunProperties41.Append(bold38);
            paragraphMarkRunProperties41.Append(boldComplexScript17);
            paragraphMarkRunProperties41.Append(italicComplexScript24);
            paragraphMarkRunProperties41.Append(nfontSize72);
            paragraphMarkRunProperties41.Append(nfontSizeComplexScript67);

            paragraphProperties41.Append(tabs8);
            paragraphProperties41.Append(spacingBetweenLines40);
            paragraphProperties41.Append(indentation21);
            paragraphProperties41.Append(justification20);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run44 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts nrunFonts85 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript18 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript25 = new ItalicComplexScript();
            FontSize nfontSize73 = new FontSize() { Val = "26" };
            FontSizeComplexScript nfontSizeComplexScript68 = new FontSizeComplexScript() { Val = "26" };

            runProperties44.Append(nrunFonts85);
            runProperties44.Append(boldComplexScript18);
            runProperties44.Append(italicComplexScript25);
            runProperties44.Append(nfontSize73);
            runProperties44.Append(nfontSizeComplexScript68);
            Text text44 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text44.Text = "Quý Thầy Cô ";

            run44.Append(runProperties44);
            run44.Append(text44);

            Run nrun45 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties nrunProperties45 = new RunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript19 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript26 = new ItalicComplexScript();
            FontSize nfontSize74 = new FontSize() { Val = "26" };
            FontSizeComplexScript nfontSizeComplexScript69 = new FontSizeComplexScript() { Val = "26" };

            nrunProperties45.Append(runFonts86);
            nrunProperties45.Append(boldComplexScript19);
            nrunProperties45.Append(italicComplexScript26);
            nrunProperties45.Append(nfontSize74);
            nrunProperties45.Append(nfontSizeComplexScript69);
            Text ntext45 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            ntext45.Text = "vui lòng ";

            nrun45.Append(nrunProperties45);
            nrun45.Append(ntext45);

            Run nrun46 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties nrunProperties46 = new RunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript20 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript27 = new ItalicComplexScript();
            FontSize fontSize75 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "26" };

            nrunProperties46.Append(runFonts87);
            nrunProperties46.Append(boldComplexScript20);
            nrunProperties46.Append(italicComplexScript27);
            nrunProperties46.Append(fontSize75);
            nrunProperties46.Append(fontSizeComplexScript70);
            Text ntext46 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            ntext46.Text = "bổ sung ";

            nrun46.Append(nrunProperties46);
            nrun46.Append(ntext46);

            Run run47 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript21 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript28 = new ItalicComplexScript();
            FontSize fontSize76 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "26" };

            runProperties47.Append(runFonts88);
            runProperties47.Append(boldComplexScript21);
            runProperties47.Append(italicComplexScript28);
            runProperties47.Append(fontSize76);
            runProperties47.Append(fontSizeComplexScript71);
            Text text47 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text47.Text = "thêm ";

            run47.Append(runProperties47);
            run47.Append(text47);

            Run run48 = new Run() { RsidRunProperties = "00043D04" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript22 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript29 = new ItalicComplexScript();
            FontSize fontSize77 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "26" };

            runProperties48.Append(runFonts89);
            runProperties48.Append(boldComplexScript22);
            runProperties48.Append(italicComplexScript29);
            runProperties48.Append(fontSize77);
            runProperties48.Append(fontSizeComplexScript72);
            Text text48 = new Text();
            text48.Text = "thông tin giảng dạy khác (nếu có) chưa được nhà trường ghi nhận, bao gồm cả việc giảng dạy tại trường Phổ thông Năng khiếu và các đơn vị thành viên ĐHQG-HCM";

            run48.Append(runProperties48);
            run48.Append(text48);

            Run run49 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript23 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript30 = new ItalicComplexScript();
            FontSize fontSize78 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "26" };

            runProperties49.Append(runFonts90);
            runProperties49.Append(boldComplexScript23);
            runProperties49.Append(italicComplexScript30);
            runProperties49.Append(fontSize78);
            runProperties49.Append(fontSizeComplexScript73);
            Text text49 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text49.Text = " và mục ";

            run49.Append(runProperties49);
            run49.Append(text49);

            Run run50 = new Run() { RsidRunProperties = "00485A86", RsidRunAddition = "00485A86" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts91 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold39 = new Bold();
            ItalicComplexScript italicComplexScript31 = new ItalicComplexScript();
            FontSize fontSize79 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "26" };

            runProperties50.Append(runFonts91);
            runProperties50.Append(bold39);
            runProperties50.Append(italicComplexScript31);
            runProperties50.Append(fontSize79);
            runProperties50.Append(fontSizeComplexScript74);
            Text text50 = new Text();
            text50.Text = "(b) Thông tin giảng dạy khác";

            run50.Append(runProperties50);
            run50.Append(text50);

            Run run51 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript24 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript32 = new ItalicComplexScript();
            FontSize fontSize80 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "26" };

            runProperties51.Append(runFonts92);
            runProperties51.Append(boldComplexScript24);
            runProperties51.Append(italicComplexScript32);
            runProperties51.Append(fontSize80);
            runProperties51.Append(fontSizeComplexScript75);
            Text text51 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text51.Text = ", ";

            run51.Append(runProperties51);
            run51.Append(text51);

            Run run52 = new Run() { RsidRunProperties = "00043D04", RsidRunAddition = "00485A86" };

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            ItalicComplexScript italicComplexScript33 = new ItalicComplexScript();
            FontSize fontSize81 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "26" };

            runProperties52.Append(runFonts93);
            runProperties52.Append(italicComplexScript33);
            runProperties52.Append(fontSize81);
            runProperties52.Append(fontSizeComplexScript76);
            Text text52 = new Text();
            text52.Text = "trong phần";

            run52.Append(runProperties52);
            run52.Append(text52);

            Run run53 = new Run() { RsidRunProperties = "00043D04", RsidRunAddition = "00485A86" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts94 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold40 = new Bold();
            BoldComplexScript boldComplexScript25 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript34 = new ItalicComplexScript();
            FontSize fontSize82 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "26" };

            runProperties53.Append(runFonts94);
            runProperties53.Append(bold40);
            runProperties53.Append(boldComplexScript25);
            runProperties53.Append(italicComplexScript34);
            runProperties53.Append(fontSize82);
            runProperties53.Append(fontSizeComplexScript77);
            Text text53 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text53.Text = " Giảng dạy lý thuyết, bài tập và thực hành";

            run53.Append(runProperties53);
            run53.Append(text53);

            Run run54 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold41 = new Bold();
            BoldComplexScript boldComplexScript26 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript35 = new ItalicComplexScript();
            FontSize fontSize83 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "26" };

            runProperties54.Append(runFonts95);
            runProperties54.Append(bold41);
            runProperties54.Append(boldComplexScript26);
            runProperties54.Append(italicComplexScript35);
            runProperties54.Append(fontSize83);
            runProperties54.Append(fontSizeComplexScript78);
            Text text54 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text54.Text = " ";

            run54.Append(runProperties54);
            run54.Append(text54);

            Run run55 = new Run() { RsidRunProperties = "00485A86", RsidRunAddition = "00485A86" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            ItalicComplexScript italicComplexScript36 = new ItalicComplexScript();
            FontSize fontSize84 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "26" };

            runProperties55.Append(runFonts96);
            runProperties55.Append(italicComplexScript36);
            runProperties55.Append(fontSize84);
            runProperties55.Append(fontSizeComplexScript79);
            Text text55 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text55.Text = "của ";

            run55.Append(runProperties55);
            run55.Append(text55);

            Run run56 = new Run() { RsidRunProperties = "00485A86", RsidRunAddition = "00485A86" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts97 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold42 = new Bold();
            BoldComplexScript boldComplexScript27 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript37 = new ItalicComplexScript();
            FontSize fontSize85 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "26" };

            runProperties56.Append(runFonts97);
            runProperties56.Append(bold42);
            runProperties56.Append(boldComplexScript27);
            runProperties56.Append(italicComplexScript37);
            runProperties56.Append(fontSize85);
            runProperties56.Append(fontSizeComplexScript80);
            Text text56 = new Text();
            text56.Text = "Báo cáo công tác năm học";

            run56.Append(runProperties56);
            run56.Append(text56);

            Run run57 = new Run() { RsidRunAddition = "00485A86" };

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold43 = new Bold();
            BoldComplexScript boldComplexScript28 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript38 = new ItalicComplexScript();
            FontSize fontSize86 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "26" };

            runProperties57.Append(runFonts98);
            runProperties57.Append(bold43);
            runProperties57.Append(boldComplexScript28);
            runProperties57.Append(italicComplexScript38);
            runProperties57.Append(fontSize86);
            runProperties57.Append(fontSizeComplexScript81);
            Text text57 = new Text();
            text57.Text = ".";

            run57.Append(runProperties57);
            run57.Append(text57);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run44);
            paragraph41.Append(nrun45);
            paragraph41.Append(nrun46);
            paragraph41.Append(run47);
            paragraph41.Append(run48);
            paragraph41.Append(run49);
            paragraph41.Append(run50);
            paragraph41.Append(run51);
            paragraph41.Append(run52);
            paragraph41.Append(run53);
            paragraph41.Append(run54);
            paragraph41.Append(run55);
            paragraph41.Append(run56);
            paragraph41.Append(run57);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "00043D04", RsidParagraphAddition = "000F1E75", RsidParagraphProperties = "00043D04", RsidRunAdditionDefault = "000F1E75", ParagraphId = "06871F8D", TextId = "7CA45A51" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();

            Tabs tabs9 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 567 };

            tabs9.Append(tabStop9);
            SpacingBetweenLines spacingBetweenLines41 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation22 = new Indentation() { End = "-357" };
            Justification justification21 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            BoldComplexScript boldComplexScript29 = new BoldComplexScript();
            ItalicComplexScript italicComplexScript39 = new ItalicComplexScript();
            FontSize fontSize87 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "26" };

            paragraphMarkRunProperties42.Append(runFonts99);
            paragraphMarkRunProperties42.Append(boldComplexScript29);
            paragraphMarkRunProperties42.Append(italicComplexScript39);
            paragraphMarkRunProperties42.Append(fontSize87);
            paragraphMarkRunProperties42.Append(fontSizeComplexScript82);

            paragraphProperties42.Append(tabs9);
            paragraphProperties42.Append(spacingBetweenLines41);
            paragraphProperties42.Append(indentation22);
            paragraphProperties42.Append(justification21);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            paragraph42.Append(paragraphProperties42);
            Paragraph paragraph43 = new Paragraph() { RsidParagraphAddition = "002F69F1", RsidRunAdditionDefault = "002F69F1", ParagraphId = "21E3E43F", TextId = "77777777" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "002F69F1", RsidSect = "00043D04" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U, Code = (UInt16Value)9U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(table1);
            body1.Append(paragraph7);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(paragraph10);
            body1.Append(paragraph11);
            body1.Append(paragraph12);
            body1.Append(paragraph13);
            body1.Append(paragraph14);
            body1.Append(table2);
            body1.Append(paragraph40);
            body1.Append(paragraph41);
            body1.Append(paragraph42);
            body1.Append(paragraph43);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            webSettings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            webSettings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            webSettings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            webSettings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            webSettings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            settings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            settings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            settings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting6 = new CompatibilitySetting() { Name = new EnumValue<CompatSettingNameValues>() { InnerText = "useWord2013TrackBottomHyphenation" }, Uri = "http://schemas.microsoft.com/office/word", Val = "0" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);
            compatibility1.Append(compatibilitySetting6);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "000F1E75" };
            Rsid rsid1 = new Rsid() { Val = "00043D04" };
            Rsid rsid2 = new Rsid() { Val = "00044F7E" };
            Rsid rsid3 = new Rsid() { Val = "000453C7" };
            Rsid rsid4 = new Rsid() { Val = "000C47F9" };
            Rsid rsid5 = new Rsid() { Val = "000F1E75" };
            Rsid rsid6 = new Rsid() { Val = "002A757A" };
            Rsid rsid7 = new Rsid() { Val = "002F69F1" };
            Rsid rsid8 = new Rsid() { Val = "004062F9" };
            Rsid rsid9 = new Rsid() { Val = "004632DA" };
            Rsid rsid10 = new Rsid() { Val = "00485A86" };
            Rsid rsid11 = new Rsid() { Val = "00493833" };
            Rsid rsid12 = new Rsid() { Val = "006A25F1" };
            Rsid rsid13 = new Rsid() { Val = "00756842" };
            Rsid rsid14 = new Rsid() { Val = "00770EDB" };
            Rsid rsid15 = new Rsid() { Val = "007F5C2D" };
            Rsid rsid16 = new Rsid() { Val = "00906534" };
            Rsid rsid17 = new Rsid() { Val = "00B041CF" };
            Rsid rsid18 = new Rsid() { Val = "00B15D92" };
            Rsid rsid19 = new Rsid() { Val = "00B70709" };
            Rsid rsid20 = new Rsid() { Val = "00C35C32" };
            Rsid rsid21 = new Rsid() { Val = "00CD281F" };
            Rsid rsid22 = new Rsid() { Val = "00DC149C" };
            Rsid rsid23 = new Rsid() { Val = "00E070B8" };
            Rsid rsid24 = new Rsid() { Val = "00EE01EC" };
            Rsid rsid25 = new Rsid() { Val = "00F04297" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);
            rsids1.Append(rsid13);
            rsids1.Append(rsid14);
            rsids1.Append(rsid15);
            rsids1.Append(rsid16);
            rsids1.Append(rsid17);
            rsids1.Append(rsid18);
            rsids1.Append(rsid19);
            rsids1.Append(rsid20);
            rsids1.Append(rsid21);
            rsids1.Append(rsid22);
            rsids1.Append(rsid23);
            rsids1.Append(rsid24);
            rsids1.Append(rsid25);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "32E817A6" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{BA1D4626-4D74-4BFB-86E8-07D2A640162D}" };

            settings1.Append(zoom1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            styles1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts100 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize88 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts100);
            runPropertiesBaseStyle1.Append(fontSize88);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript83);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines42 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines42);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", UiPriority = 0, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo() { Name = "Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo() { Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo374 = new LatentStyleExceptionInfo() { Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo375 = new LatentStyleExceptionInfo() { Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo376 = new LatentStyleExceptionInfo() { Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);
            latentStyles1.Append(latentStyleExceptionInfo372);
            latentStyles1.Append(latentStyleExceptionInfo373);
            latentStyles1.Append(latentStyleExceptionInfo374);
            latentStyles1.Append(latentStyleExceptionInfo375);
            latentStyles1.Append(latentStyleExceptionInfo376);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid26 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines43 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines43);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts101 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            FontSize fontSize89 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(runFonts101);
            styleRunProperties1.Append(fontSize89);
            styleRunProperties1.Append(fontSizeComplexScript84);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid26);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            StyleName styleName2 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading2Char" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid27 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines44 = new SpacingBetweenLines() { Before = "100", BeforeAutoSpacing = true, After = "100", AfterAutoSpacing = true, Line = "360", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation23 = new Indentation() { End = "-323" };
            Justification justification22 = new Justification() { Val = JustificationValues.Both };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(spacingBetweenLines44);
            styleParagraphProperties2.Append(indentation23);
            styleParagraphProperties2.Append(justification22);
            styleParagraphProperties2.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            Bold bold44 = new Bold();
            BoldComplexScript boldComplexScript30 = new BoldComplexScript();
            FontSize fontSize90 = new FontSize() { Val = "26" };

            styleRunProperties2.Append(runFonts102);
            styleRunProperties2.Append(bold44);
            styleRunProperties2.Append(boldComplexScript30);
            styleRunProperties2.Append(fontSize90);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(primaryStyle2);
            style2.Append(rsid27);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading3" };
            StyleName styleName3 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Heading3Char" };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid28 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            Indentation indentation24 = new Indentation() { End = "-235" };
            Justification justification23 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties3.Append(keepNext2);
            styleParagraphProperties3.Append(indentation24);
            styleParagraphProperties3.Append(justification23);
            styleParagraphProperties3.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            Bold bold45 = new Bold();
            Italic italic12 = new Italic();
            FontSize fontSize91 = new FontSize() { Val = "22" };

            styleRunProperties3.Append(bold45);
            styleRunProperties3.Append(italic12);
            styleRunProperties3.Append(fontSize91);

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(nextParagraphStyle2);
            style3.Append(linkedStyle2);
            style3.Append(primaryStyle3);
            style3.Append(rsid28);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading4" };
            StyleName styleName4 = new StyleName() { Val = "heading 4" };
            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "Heading4Char" };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid29 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            KeepNext keepNext3 = new KeepNext();
            Indentation indentation25 = new Indentation() { Start = "-176", End = "-134" };
            Justification justification24 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 3 };

            styleParagraphProperties4.Append(keepNext3);
            styleParagraphProperties4.Append(indentation25);
            styleParagraphProperties4.Append(justification24);
            styleParagraphProperties4.Append(outlineLevel3);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            Bold bold46 = new Bold();
            Italic italic13 = new Italic();
            FontSize fontSize92 = new FontSize() { Val = "22" };

            styleRunProperties4.Append(bold46);
            styleRunProperties4.Append(italic13);
            styleRunProperties4.Append(fontSize92);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(nextParagraphStyle3);
            style4.Append(linkedStyle3);
            style4.Append(primaryStyle4);
            style4.Append(rsid29);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading6" };
            StyleName styleName5 = new StyleName() { Val = "heading 6" };
            BasedOn basedOn4 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "Heading6Char" };
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid30 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            KeepNext keepNext4 = new KeepNext();
            Indentation indentation26 = new Indentation() { End = "-135" };
            Justification justification25 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 5 };

            styleParagraphProperties5.Append(keepNext4);
            styleParagraphProperties5.Append(indentation26);
            styleParagraphProperties5.Append(justification25);
            styleParagraphProperties5.Append(outlineLevel4);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            Bold bold47 = new Bold();
            Italic italic14 = new Italic();
            FontSize fontSize93 = new FontSize() { Val = "22" };

            styleRunProperties5.Append(bold47);
            styleRunProperties5.Append(italic14);
            styleRunProperties5.Append(fontSize93);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(nextParagraphStyle4);
            style5.Append(linkedStyle4);
            style5.Append(primaryStyle5);
            style5.Append(rsid30);
            style5.Append(styleParagraphProperties5);
            style5.Append(styleRunProperties5);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading7" };
            StyleName styleName6 = new StyleName() { Val = "heading 7" };
            BasedOn basedOn5 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() { Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "Heading7Char" };
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid31 = new Rsid() { Val = "000F1E75" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            KeepNext keepNext5 = new KeepNext();
            Indentation indentation27 = new Indentation() { End = "-109" };
            Justification justification26 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel5 = new OutlineLevel() { Val = 6 };

            styleParagraphProperties6.Append(keepNext5);
            styleParagraphProperties6.Append(indentation27);
            styleParagraphProperties6.Append(justification26);
            styleParagraphProperties6.Append(outlineLevel5);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            Bold bold48 = new Bold();
            Italic italic15 = new Italic();
            FontSize fontSize94 = new FontSize() { Val = "22" };

            styleRunProperties6.Append(bold48);
            styleRunProperties6.Append(italic15);
            styleRunProperties6.Append(fontSize94);

            style6.Append(styleName6);
            style6.Append(basedOn5);
            style6.Append(nextParagraphStyle5);
            style6.Append(linkedStyle5);
            style6.Append(primaryStyle6);
            style6.Append(rsid31);
            style6.Append(styleParagraphProperties6);
            style6.Append(styleRunProperties6);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName7 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style7.Append(styleName7);
            style7.Append(uIPriority1);
            style7.Append(semiHidden1);
            style7.Append(unhideWhenUsed1);

            Style style8 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName8 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style8.Append(styleName8);
            style8.Append(uIPriority2);
            style8.Append(semiHidden2);
            style8.Append(unhideWhenUsed2);
            style8.Append(styleTableProperties1);

            Style style9 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName9 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style9.Append(styleName9);
            style9.Append(uIPriority3);
            style9.Append(semiHidden3);
            style9.Append(unhideWhenUsed3);

            Style style10 = new Style() { Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "Heading 2 Char" };
            BasedOn basedOn6 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "Heading2" };
            Rsid rsid32 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold49 = new Bold();
            BoldComplexScript boldComplexScript31 = new BoldComplexScript();
            FontSize fontSize95 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties7.Append(runFonts103);
            styleRunProperties7.Append(bold49);
            styleRunProperties7.Append(boldComplexScript31);
            styleRunProperties7.Append(fontSize95);
            styleRunProperties7.Append(fontSizeComplexScript85);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(linkedStyle6);
            style10.Append(rsid32);
            style10.Append(styleRunProperties7);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "Heading3Char", CustomStyle = true };
            StyleName styleName11 = new StyleName() { Val = "Heading 3 Char" };
            BasedOn basedOn7 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "Heading3" };
            Rsid rsid33 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold50 = new Bold();
            Italic italic16 = new Italic();
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties8.Append(runFonts104);
            styleRunProperties8.Append(bold50);
            styleRunProperties8.Append(italic16);
            styleRunProperties8.Append(fontSizeComplexScript86);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(linkedStyle7);
            style11.Append(rsid33);
            style11.Append(styleRunProperties8);

            Style style12 = new Style() { Type = StyleValues.Character, StyleId = "Heading4Char", CustomStyle = true };
            StyleName styleName12 = new StyleName() { Val = "Heading 4 Char" };
            BasedOn basedOn8 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "Heading4" };
            Rsid rsid34 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold51 = new Bold();
            Italic italic17 = new Italic();
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties9.Append(runFonts105);
            styleRunProperties9.Append(bold51);
            styleRunProperties9.Append(italic17);
            styleRunProperties9.Append(fontSizeComplexScript87);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(linkedStyle8);
            style12.Append(rsid34);
            style12.Append(styleRunProperties9);

            Style style13 = new Style() { Type = StyleValues.Character, StyleId = "Heading6Char", CustomStyle = true };
            StyleName styleName13 = new StyleName() { Val = "Heading 6 Char" };
            BasedOn basedOn9 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "Heading6" };
            Rsid rsid35 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold52 = new Bold();
            Italic italic18 = new Italic();
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties10.Append(runFonts106);
            styleRunProperties10.Append(bold52);
            styleRunProperties10.Append(italic18);
            styleRunProperties10.Append(fontSizeComplexScript88);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(linkedStyle9);
            style13.Append(rsid35);
            style13.Append(styleRunProperties10);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "Heading7Char", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "Heading 7 Char" };
            BasedOn basedOn10 = new BasedOn() { Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "Heading7" };
            Rsid rsid36 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold53 = new Bold();
            Italic italic19 = new Italic();
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties11.Append(runFonts107);
            styleRunProperties11.Append(bold53);
            styleRunProperties11.Append(italic19);
            styleRunProperties11.Append(fontSizeComplexScript89);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(linkedStyle10);
            style14.Append(rsid36);
            style14.Append(styleRunProperties11);

            Style style15 = new Style() { Type = StyleValues.Character, StyleId = "Hyperlink" };
            StyleName styleName15 = new StyleName() { Val = "Hyperlink" };
            Rsid rsid37 = new Rsid() { Val = "000F1E75" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            Color color1 = new Color() { Val = "0000FF" };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties12.Append(color1);
            styleRunProperties12.Append(underline1);

            style15.Append(styleName15);
            style15.Append(rsid37);
            style15.Append(styleRunProperties12);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            fonts1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            fonts1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Times New Roman" };
            AltName altName1 = new AltName() { Val = "Calibri" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00000001", CodePageSignature1 = "00000000" };

            font3.Append(altName1);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "A0002AEF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number3);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            fontTablePart1.Fonts = fonts1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Minh-Triet TRAN";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "4";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-06-29T04:01:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-06-29T04:22:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Minh-Triet TRAN";
        }


    }
}
