using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace GenerateDocxDemo
{

    /// <summary>
    /// This class is responsible for building up a document that displays random numbers as word art.
    /// </summary>
    public class ExtraDocument
    {
        /// <summary>
        /// Empty Constructor
        /// </summary>
        public ExtraDocument()
        {
        }

        /// <summary>
        /// Builds the document
        /// </summary>
        /// <returns>A byte array containing the content of the document.</returns>
        public byte[] GenerateDocument(int numPeople, int numGiveaways)
        {
            //Split each random number out.
            string[] randNumbers = GetRandomNumbers(numPeople, numGiveaways).Split(',');

            //maintain a list of paragraphs to later be appended to the body of our document.
            List<Paragraph> randPars = new List<Paragraph>();

            //Loop through each random number and create a word art item.
            for (int i = 0; i <= randNumbers.Count() - 1; i++)
            {
                //The code below is building up the word art that will be appended to our document.
                Picture picture1 = new Picture();

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t175", CoordinateSize = "21600,21600", OptionalNumber = 175, Adjustment = "3086", EdgePath = "m,qy10800@0,21600,m0@1qy10800,21600,21600@1e" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "val #0" };
                V.Formula formula2 = new V.Formula() { Equation = "sum 21600 0 #0" };
                V.Formula formula3 = new V.Formula() { Equation = "prod @1 1 2" };
                V.Formula formula4 = new V.Formula() { Equation = "sum @2 10800 0" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "10800,@0;0,@2;10800,21600;21600,@2", ConnectAngles = "270,180,90,0" };
                V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

                V.ShapeHandles handles1 = new V.ShapeHandles();
                V.ShapeHandle handle1 = new V.ShapeHandle() { Position = "center,#0", YRange = "0,7200" };

                handles1.Append(handle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

                shapetype1.Append(formulas1);
                shapetype1.Append(path1);
                shapetype1.Append(textPath1);
                shapetype1.Append(handles1);
                shapetype1.Append(lock1);

                V.Shape shape1 = new V.Shape() { Id = "_x0000_i1025", Style = "width:159pt;height:87pt", FillColor = "#d6e3bc [1302]", StrokeColor = "#009", StrokeWeight = "4.5pt", Type = "#_x0000_t175", Adjustment = ",10800" };
                V.Shadow shadow1 = new V.Shadow() { On = true, Color = "#009", Offset = "7pt,-7pt" };
                V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Impact\";v-text-spacing:52429f;v-text-kern:t", FitPath = true, Trim = true, String = randNumbers[i] };

                shape1.Append(shadow1);
                shape1.Append(textPath2);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                //Add to our word art paragraph list.
                randPars.Add(new Paragraph(new Run(picture1)));
            }

            //Create the document that will be used to display the random numbers.
            byte[] docBytes;
            using (MemoryStream memStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    mainPart.Document.Body = new Body();

                    //Add each word art paragraph to the body of the document.
                    foreach(Paragraph par in randPars)
                    {
                        mainPart.Document.Body.Append(par);
                    }
                }

                //Convert the document to a byte array so that it can be written to the response.
                docBytes = memStream.ToArray();
            }

            return docBytes;
        }

        /// <summary>
        /// Gets a list of random numbers and returns them as a comma delimited string.
        /// </summary>
        private string GetRandomNumbers(int numPeople, int numGiveaways)
        {
            List<int> randoms = new List<int>();

            Random rand = new Random();

            for (int i = 0; i <= numGiveaways; i++)
            {
                int j = rand.Next(1, numPeople);

                while (randoms.Contains(j))
                {
                    j = rand.Next(1, numPeople);
                }

                randoms.Add(j);
            }

            return String.Join(",", randoms.ConvertAll(i => i.ToString()).ToArray());
        }
    }
}
