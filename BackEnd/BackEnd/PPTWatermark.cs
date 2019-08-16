using System;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace BackEnd
{
    public class PPTWatermark
    {
        public static void ChangeSlideMasterPart(SlideMasterPart slideMasterPart1)
        {
            //Insert code here that adds text box to master
            SlideMaster slideMaster1 = slideMasterPart1.SlideMaster;

            CommonSlideData commonSlideData1 = slideMaster1.GetFirstChild<CommonSlideData>();

            ShapeTree shapeTree1 = commonSlideData1.GetFirstChild<ShapeTree>();

            Shape shape1 = new Shape(); //create shape for text box

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "TextBox 2" };
            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();
            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3EA4C13C-15CC-4826-AFD0-56FC7F217884}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);
            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);
            nonVisualDrawingProperties5.Append(nonVisualDrawingPropertiesExtensionList1);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties() { UserDrawn = true };

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties5);

            ShapeProperties shapeProperties3 = new ShapeProperties(); //shape properties, can edit here

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset3 = new A.Offset() { X = 5001861L, Y = 5943491L };
            A.Extents extents3 = new A.Extents() { Cx = 3608739L, Cy = 646331L };

            transform2D1.Append(offset3);
            transform2D1.Append(extents3);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill2 = new A.NoFill();

            shapeProperties3.Append(transform2D1);
            shapeProperties3.Append(presetGeometry1);
            shapeProperties3.Append(noFill2);

            TextBody textBody3 = new TextBody(); //where the text is going to be added

            A.BodyProperties bodyProperties3 = new A.BodyProperties()
            {
                Wrap = A.TextWrappingValues.Square,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Center
            };
            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties3.Append(shapeAutoFit1);
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph(); //first paragraph, full name will go here
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", Dirty = false };

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 40000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 60000 };
            A.Alpha alpha1 = new A.Alpha() { Val = 79000 };


            schemeColor15.Append(luminanceModulation1);
            schemeColor15.Append(luminanceOffset1);
            schemeColor15.Append(alpha1);

            solidFill6.Append(schemeColor15);

            runProperties1.Append(solidFill6);
            A.Text text1 = new A.Text
            {
                Text = "Microsoft Corporation" //full name 
            };

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph3.Append(paragraphProperties1);
            paragraph3.Append(run1);

            A.Paragraph paragraph4 = new A.Paragraph(); //second paragraph contains the dcxhelp@microsoft.com email
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", Dirty = false };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 40000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 60000 };
            A.Alpha alpha2 = new A.Alpha() { Val = 79000 };


            schemeColor16.Append(luminanceModulation2);
            schemeColor16.Append(luminanceOffset2);
            schemeColor16.Append(alpha2);

            solidFill7.Append(schemeColor16);

            runProperties2.Append(solidFill7);
            A.Text text2 = new A.Text
            {
                Text = "dcxhelp@microsoft.com"
            };

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph4.Append(paragraphProperties2);
            paragraph4.Append(run2);

            //Microsoft Corporate Address, edit if incorrect



            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);
            textBody3.Append(paragraph4);


            shape1.Append(nonVisualShapeProperties3);
            shape1.Append(shapeProperties3);
            shape1.Append(textBody3);

            shapeTree1.Append(shape1);

        }
    }
}
