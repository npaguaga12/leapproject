using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using HorizontalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.HorizontalAnchorValues;
using Lock = DocumentFormat.OpenXml.Vml.Office.Lock;
using Path = DocumentFormat.OpenXml.Vml.Path;
using VerticalAnchorValues = DocumentFormat.OpenXml.Vml.Wordprocessing.VerticalAnchorValues;

namespace BackEnd
{
    public class DocWatermark
    {
        static Header MakeHeader()
        {
            var header = new Header();
            var paragraph = new Paragraph();
            var run = new Run();
            var text = new Text();
            text.Text = "";
            run.Append(text);
            paragraph.Append(run);
            header.Append(paragraph);
            return header;
        }


        public static void AddWaterMark(WordprocessingDocument doc)
        {

            if (doc.MainDocumentPart.HeaderParts.Count() == 0)
            {
                doc.MainDocumentPart.DeleteParts(doc.MainDocumentPart.HeaderParts);
                var newHeaderPart = doc.MainDocumentPart.AddNewPart<HeaderPart>();
                var rId = doc.MainDocumentPart.GetIdOfPart(newHeaderPart);
                var headerRef = new HeaderReference();
                headerRef.Id = rId;
                var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
                if (sectionProps == null)
                {
                    sectionProps = new SectionProperties();
                    doc.MainDocumentPart.Document.Body.Append(sectionProps);
                }
                sectionProps.RemoveAllChildren<HeaderReference>();
                sectionProps.Append(headerRef);

                newHeaderPart.Header = MakeHeader();
                newHeaderPart.Header.Save();
            }
            foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
            {


                SdtBlock sdtBlock1 = new SdtBlock();
                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 87908844 };
                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();

                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);
                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);


                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();
                Paragraph paragraph2 = new Paragraph()
                {
                    RsidParagraphAddition = "00656E18",
                    RsidRunAdditionDefault = "00656E18"
                };
                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };
                paragraphProperties2.Append(paragraphStyleId2);


                Run run1 = new Run();
                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();
                Languages languages1 = new Languages() { EastAsia = "zh-TW" };
                runProperties1.Append(noProof1);
                runProperties1.Append(languages1);


                Picture picture1 = new Picture();

                Shapetype shapetype1 = new Shapetype()
                {
                    Id = "_x0000_t136",
                    CoordinateSize = "21600,21600",
                    OptionalNumber = 136,
                    Adjustment = "10800",
                    EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
                };

                Formulas formulas1 = new Formulas();
                Formula formula1 = new Formula() { Equation = "sum #0 0 10800" };
                Formula formula2 = new Formula() { Equation = "prod #0 2 1" };
                Formula formula3 = new Formula() { Equation = "sum 21600 0 @1" };
                Formula formula4 = new Formula() { Equation = "sum 0 0 @2" };
                Formula formula5 = new Formula() { Equation = "sum 21600 0 @3" };
                Formula formula6 = new Formula() { Equation = "if @0 @3 0" };
                Formula formula7 = new Formula() { Equation = "if @0 21600 @1" };
                Formula formula8 = new Formula() { Equation = "if @0 0 @2" };
                Formula formula9 = new Formula() { Equation = "if @0 @4 21600" };
                Formula formula10 = new Formula() { Equation = "mid @5 @6" };
                Formula formula11 = new Formula() { Equation = "mid @8 @5" };
                Formula formula12 = new Formula() { Equation = "mid @7 @8" };
                Formula formula13 = new Formula() { Equation = "mid @6 @7" };
                Formula formula14 = new Formula() { Equation = "sum @6 0 @5" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                formulas1.Append(formula11);
                formulas1.Append(formula12);
                formulas1.Append(formula13);
                formulas1.Append(formula14);


                Path path1 = new Path()
                {
                    AllowTextPath = true,
                    ConnectionPointType = ConnectValues.Custom,
                    ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800",
                    ConnectAngles = "270,180,90,0"
                };
                TextPath textPath1 = new TextPath()
                {
                    On = true,
                    FitShape = true
                };

                ShapeHandles shapeHandles1 = new ShapeHandles();
                ShapeHandle shapeHandle1 = new ShapeHandle()
                {
                    Position = "#0,bottomRight",
                    XRange = "6629,14971"
                };
                shapeHandles1.Append(shapeHandle1);


                Lock lock1 = new Lock()
                {
                    Extension = ExtensionHandlingBehaviorValues.Edit,
                    TextLock = true,
                    ShapeType = true
                };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(shapeHandles1);
                shapetype1.Append(lock1);


                Shape shape1 = new Shape()
                {
                    Id = "PowerPlusWaterMarkObject357476642",
                    Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:280pt;height:175pt;rotation:360;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                    OptionalString = "_x0000_s2049",
                    AllowInCell = false,
                    FillColor = "blue",
                    Stroked = false,
                    Type = "#_x0000_t136"
                };


                Fill fill1 = new Fill() { Opacity = ".9" };

                var currentTime = DateTime.Now;
                string strTime = currentTime.ToString("MM/dd/yyyy HH:mm");
                string txtMark = $"Microsoft Coroporation\nOne Microsoft Way\nRedmond, WA 98052\ndcxhelp@microsoft.com".Replace("\n", Environment.NewLine);
                TextPath textPath2 = new TextPath()
                {
                    Style = "font-family:\"Calibri\";font-size:1pt",
                    String = txtMark
                };

                TextWrap textWrap1 = new TextWrap()
                {
                    AnchorX = HorizontalAnchorValues.Margin,
                    AnchorY = VerticalAnchorValues.Margin
                };

                shape1.Append(fill1);
                shape1.Append(textPath2);
                shape1.Append(textWrap1);
                picture1.Append(shapetype1);
                picture1.Append(shape1);
                run1.Append(runProperties1);
                run1.Append(picture1);
                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run1);
                sdtContentBlock1.Append(paragraph2);
                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                headerPart.Header.Append(sdtBlock1);
            }
        }
    }
}
