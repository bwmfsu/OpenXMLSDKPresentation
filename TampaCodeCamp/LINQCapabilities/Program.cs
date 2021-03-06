﻿using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Word = Microsoft.Office.Interop.Word;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace LINQCapabilities
{
    class Program
    {
        private static WordprocessingDocument concertDoc;

        static void Main(string[] args)
        {
            string originalDoc = string.Format(@"{0}\{1}", Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName, "TheBestConcertEver.docx");
            string modifiedDoc = string.Format(@"{0}\{1}", Directory.GetCurrentDirectory(), "TheBestConcertEver[PendingReview].docx");

            if (File.Exists(modifiedDoc))
            {
                File.Delete(modifiedDoc);
            }

            File.Copy(originalDoc, modifiedDoc);            

            //Open our document and begin manipulating it.
            using (concertDoc = WordprocessingDocument.Open(modifiedDoc, true))
            {
                //Billy Joel won't be able to play until 3:30 now.  Discussed with Boston and they are fine with switching time slots.
                //Find the time slot and change it to 3:30
                ChangeBandTime("Billy Joel", "1:30 PM", "3:30 PM");
                ChangeBandTime("Boston", "3:30 PM", "1:30 PM");

                //Mark the document as confidential because we don't want anyone knowing this information until the day of the concert.
                AddWatermark();
            }

            //Open the document in Word.
            OpenDocument(modifiedDoc);

        }

        /// <summary>
        /// Finds the set time for a band and changes it to the new time.
        /// </summary>
        /// <param name="bandName">The band to change the time for.</param>
        /// <param name="oldTime">The old time that needs to be changed.</param>
        /// <param name="newTime">The new time slot.</param>
        private static void ChangeBandTime(string bandName, string oldTime, string newTime)
        {
            //Get the Text containing the band whose time we want to change.
            Text band = concertDoc.MainDocumentPart.Document.Descendants<Text>().Where(text => text.InnerText == bandName).SingleOrDefault();

            //Get the time slot for the band and update it to the new time.
            TableRow bandRow = band.Ancestors<TableRow>().SingleOrDefault();
            Text newTimeSlot = bandRow.Descendants<Text>().Where(timeSlot => timeSlot.Text == oldTime).SingleOrDefault();
            newTimeSlot.Text = newTime;

            //Indicate that the text has changed by setting the text color to green.
            Run timeSlotRun = newTimeSlot.Parent as Run;
            timeSlotRun.RunProperties.Append(new Color() { Val = "green" });
        }

        /// <summary>
        /// Sets the a watermark of "Pending Approval" for the document to show that a change has been made, but still needs to be reviewed by the 
        /// concert organizer.
        /// </summary>
        /// <remarks>
        /// A lot of this code came directly out of code generated by the Open XML SDK tool for a document that had a watermark.
        /// </remarks>
        private static void AddWatermark()
        {
            Header waterMarkHeader = new Header();

            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtId sdtId1 = new SdtId() { Val = 181394437 };

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Watermarks" };
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00380641", RsidRunAdditionDefault = "00380641" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();
            Languages languages1 = new Languages() { EastAsia = "zh-TW" };

            runProperties1.Append(noProof1);
            runProperties1.Append(languages1);

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
            V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
            V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
            V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
            V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
            V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
            V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
            V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
            V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
            V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
            V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
            V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

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
            V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
            V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

            V.ShapeHandles handles1 = new V.ShapeHandles();
            V.ShapeHandle handle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

            handles1.Append(handle1);
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(textPath1);
            shapetype1.Append(handles1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape() { Id = "PowerPlusWaterMarkObject357476642", Style = "position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s3073", AllowInCell = false, FillColor = "silver", Stroked = false, Type = "#_x0000_t136" };
            V.Fill fill1 = new V.Fill() { Opacity = ".5" };
            V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Calibri\";font-size:1pt", String = "PENDING APPROVAL" };
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

            shape1.Append(fill1);
            shape1.Append(textPath2);
            shape1.Append(textWrap1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run1.Append(runProperties1);
            run1.Append(picture1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);

            waterMarkHeader.Append(sdtBlock1);

            HeaderPart docHeaderPart = concertDoc.MainDocumentPart.AddNewPart<HeaderPart>();
            docHeaderPart.Header = waterMarkHeader;

           SectionProperties sectProps = concertDoc.MainDocumentPart.Document.Descendants<SectionProperties>().SingleOrDefault();
           HeaderReference headerRef = new HeaderReference();
           headerRef.Id = concertDoc.MainDocumentPart.GetIdOfPart(docHeaderPart);
           headerRef.Type = HeaderFooterValues.Default;
           sectProps.Append(headerRef);
        }

        /// <summary>
        /// Open the copied document using Word.Interop.
        /// </summary>
        /// <param name="copiedDoc">The document to open.</param>
        private static void OpenDocument(string copiedDoc)
        {
            object oFileName = copiedDoc;
            object oMissing = System.Reflection.Missing.Value;
            object oTrue = true;

            Word.ApplicationClass wordApp = new Word.ApplicationClass();
            wordApp.Visible = true;
            wordApp.Documents.Open(ref oFileName, ref oMissing, ref oTrue, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oTrue, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
        }
    }
}
