using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;

using System.IO;

namespace document_generation_functions
{
    public class functions
    {
        public string generate_doc(string template_path,string output_path)
        {
            string templatePath = template_path;
            string outputPath = output_path;


            File.Copy(templatePath, outputPath, true);

            // Create a new document
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;

                // You can now write to the document body
                var body = mainPart.Document.Body;

                //MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                //mainPart.Document = new Document();
                //Body body = new Body();

                //paragraph.Append(paraProps);

                // Title
                //body.Append(new Paragraph(new Run(new Text("Project Report"))) { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = "Title" }) });

                // Add placeholder paragraphs
                body.Append(CreateMixedBoldParagraph("الرقم", "***reference number***"));

                body.Append(CreateMixedBoldParagraph("الموافق ", "***hijri date***"));
                body.Append(CreateMixedBoldParagraph("التصنيف  ", "***security classification***"));
                body.Append(new Paragraph(new Run(new Break())));

                body.Append(CreateMixedBoldParagraph("الفاضل/-    ***Receiver***           المحترم", ""));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateMixedBoldParagraph("تحية طيبة،،،", ""));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateUnderlinedCenteredParagraph("الموضوع‏:‏***subject***"));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateStyledParagraph("أكتب نص الرسالة هنا...", false, false, false, "FF0000"));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateMixedBoldParagraph("وتقبلوا فائق الاحترام و التقدير", ""));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateMixedBoldParagraph("***User***", ""));
                body.Append(CreateMixedBoldParagraph("***Position***", ""));
                body.Append(new Paragraph(new Run(new Break())));
                body.Append(CreateMixedBoldParagraph("نسخة الى", "***copy to***"));

                var imagePart = mainPart.AddImagePart(ImagePartType.Png); // or Png, Bmp etc.

                var stamp_type = "ceo";

                if (stamp_type == "ceo")
                {
                    using (var stream = new FileStream("ceologo.png", FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                }
                else
                {
                    using (var stream = new FileStream("obblogo.png", FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                }

                string imageRelId = mainPart.GetIdOfPart(imagePart);

                // 2. Define the image paragraph (fixed size image: ~5cm x 3cm)
                var imagePara = new Paragraph(new Run(CreateImageElement(imageRelId, 1800000, 1000000)));

                // 3. Append it to the document body
                body.AppendChild(imagePara);


                //mainPart.Document.Append(body);
                mainPart.Document.Save();
            }

            Console.WriteLine("Word document created successfully.");
            return ("document_generated");
        }

        static Paragraph CreateParagraph(string text)
        {
            return new Paragraph(new Run(new Text(text)));
        }

        static Paragraph CreateHeading(string text)
        {
            return new Paragraph(
                new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }),
                new Run(new Text(text))
            );
        }

        static Paragraph CreateMixedBoldParagraph(string boldText, string normalText)
        {
            // Normal text run


            // Set RTL properties
            RunProperties R2LProps = new RunProperties();
            R2LProps.RightToLeftText = new RightToLeftText(); // RTL for run

            Run normalRun = new Run(R2LProps, new Text(normalText));

            // Bold text run
            RunProperties boldProps = new RunProperties(new Bold());
            boldProps.RightToLeftText = new RightToLeftText(); // RTL for run

            //run.Append(runProps);
            Run boldRun = new Run(boldProps, new Text(boldText));
            var colonRun = new Run(
                new RunProperties(new RightToLeftText()),
                new Text("\u200F:\u200F") // U+200F is Right-to-Left Mark
            );
            Paragraph prgph = new Paragraph();
            ParagraphProperties paraProps = new ParagraphProperties();
            paraProps.BiDi = new BiDi(); // Sets paragraph as bidirectional (RTL)
                                         //paraProps.Justification = new Justification() { Val = JustificationValues.Left };
            prgph.Append(paraProps);
            if (normalText == "")
            {
                prgph.Append(boldRun, normalRun);
                //= new Paragraph(boldRun, normalRun);
            }
            else
            {
                prgph.Append(boldRun, colonRun, normalRun);
                //Paragraph prgph = new Paragraph(boldRun, colonRun, normalRun);
            }

    ;

            // Combine runs into one paragraph
            return prgph;
        }

        static Paragraph CreateUnderlinedCenteredParagraph(string text)
        {
            // Set run properties: underline
            RunProperties runProps = new RunProperties();
            RunProperties boldProps = new RunProperties(new Bold());


            boldProps.RightToLeftText = new RightToLeftText(); // RTL for run
            runProps.Append(new Underline() { Val = UnderlineValues.Single });

            // Create run with text
            Run run = new Run();
            run.Append(runProps);
            run.Append(boldProps);
            run.Append(new Text(text));

            // Set paragraph properties: center alignment
            ParagraphProperties paraProps = new ParagraphProperties();
            paraProps.Append(new Justification() { Val = JustificationValues.Center });

            // Combine everything into paragraph
            Paragraph paragraph = new Paragraph();
            paraProps.BiDi = new BiDi(); // Sets paragraph as bidirectional (RTL)
            paragraph.Append(paraProps);
            paragraph.Append(run);

            return paragraph;
        }

        static Paragraph CreateStyledParagraph(string text, bool bold = false, bool underline = false, bool centered = false, string hexColor = "FF0000")
        {
            // Create run properties
            RunProperties runProps = new RunProperties();
            runProps.RightToLeftText = new RightToLeftText(); // RTL for run

            if (bold)
                runProps.Append(new Bold());

            if (underline)
                runProps.Append(new Underline() { Val = UnderlineValues.Single });

            if (!string.IsNullOrEmpty(hexColor))
                runProps.Append(new Color() { Val = hexColor }); // "FF0000" = red

            // Create the run
            Run run = new Run();
            run.Append(runProps);
            run.Append(new Text(text));

            // Paragraph properties
            Paragraph paragraph = new Paragraph();

            if (centered)
            {
                ParagraphProperties paraProps = new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }

                );
                //paraProps.BiDi = new BiDi(); // Sets paragraph as bidirectional (RTL)
                paragraph.Append(paraProps);
            }
            ParagraphProperties paraProps2 = new ParagraphProperties();
            paraProps2.BiDi = new BiDi(); // Sets paragraph as bidirectional (RTL)
            paragraph.Append(paraProps2);

            paragraph.Append(run);
            return paragraph;
        }

        static Drawing CreateImageElement(string relationshipId, long cx, long cy)
        {
            return new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = cx, Cy = cy },
                    new DW.DocProperties() { Id = 1U, Name = "Image" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Image" },
                                    new PIC.NonVisualPictureDrawingProperties()
                                ),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())
                                ),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0, Y = 0 },
                                        new A.Extents() { Cx = cx, Cy = cy }
                                    ),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
            );
        }
    }
}
