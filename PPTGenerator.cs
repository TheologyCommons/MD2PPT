using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace Md_Ppt
{
    class PPTGenerator
    {
        internal static void Create(string markDownFile, string targetFile)
        {
            var paragraphs = new List<Paragraph>();
            using (StreamReader fs = new StreamReader(markDownFile))
            {
                var paragraph = new Paragraph();
                while (!fs.EndOfStream)
                {
                    var line = fs.ReadLine();
                    if (line.All(c => c == 13))
                    {
                        paragraphs.Add(paragraph);
                        paragraph = new Paragraph();
                    }
                    else if (!string.IsNullOrWhiteSpace(line))
                    {
                        paragraph.Add(line);
                    }
                }
            }
            var para = new Paragraph(){System.IO.Path.GetFileNameWithoutExtension(markDownFile) };
            var outlineParagraph = new OutlineParagraph(para, paragraphs);
            if (outlineParagraph.Children.Count > 10)
            {
                foreach (var p in outlineParagraph.Children)
                {
                    createPowerPoint(p, targetFile);
                }
            }
            else
            {
                createPowerPoint(outlineParagraph, targetFile);
            }
        }

        private static void createPowerPoint(OutlineParagraph outlineParagraph, string targetFile)
        {
            var pptApplication = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();

            var rootPath = System.IO.Path.GetDirectoryName(targetFile);
            var targetName = System.IO.Path.GetFileNameWithoutExtension(targetFile);
            var pptFileTargetPath = System.IO.Path.Combine(rootPath, targetName) + $"\\" + outlineParagraph.Text.Trim() + ".ppt";


            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);

            var filePath = typeof(PPTGenerator).Assembly.Location;
            var path = System.IO.Path.GetDirectoryName(filePath);


            pptPresentation.ApplyTheme(path + @"\BibleStudy.thmx");
            createPowerPoint(outlineParagraph, pptPresentation);

            Console.WriteLine(pptFileTargetPath);

            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(pptFileTargetPath));

                pptPresentation.SaveAs(pptFileTargetPath);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }

            var odpOutputfile = Path.Combine(rootPath, "..\\ODP", targetName, outlineParagraph.Text.Trim()+ ".odp");



            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(odpOutputfile));
            pptPresentation.SaveAs(odpOutputfile, PpSaveAsFileType.ppSaveAsOpenDocumentPresentation);

            pptPresentation.Close();
            pptApplication.Quit();

        }

        private static void createPowerPoint(OutlineParagraph outlineParagraph, Presentation pptPresentation)
        {
            var text = outlineParagraph.Text;
            var subTexts = new List<string>();
            if (outlineParagraph.Level == 0)
                createLayoutTitle(text, subTexts, pptPresentation);
            foreach (var child in outlineParagraph.Children)
            {
                if (outlineParagraph.Level > 0) 
                    subTexts.Add(cleanUpText(child.Text));
                if (child.Children.Any())
                {
                    if (outlineParagraph.Level > 0)
                    {
                        createLayoutTitle(text, subTexts, pptPresentation);
                    }

                    createPowerPoint(child, pptPresentation);
                }
            }
            createLayoutTitle(text, subTexts, pptPresentation);

        }

        private static void createLayoutTitle(string text, List<string> subTexts, Presentation pptPresentation)
        {
            var processedText = cleanUpText(text);
            if (!String.IsNullOrWhiteSpace(text))
            {
                int newSlideNumber = (pptPresentation.Slides.Count + 1);
                var layout = PpSlideLayout.ppLayoutTitle;
                if (subTexts.Count >= 2)
                {
                    layout = PpSlideLayout.ppLayoutText;
                }
                var subTextBuilder = new StringBuilder();
                var subTextBuilder2 = new StringBuilder();

                if (subTexts.Count > 6)
                {
                    layout = PpSlideLayout.ppLayoutTwoColumnText;
                    for (int i = 0; i < subTexts.Count; i++)
                    {
                        if (i < subTexts.Count / 2)
                        {
                            subTextBuilder.AppendLine(subTexts[i]);
                        }
                        else
                        {
                            subTextBuilder2.AppendLine(subTexts[i]);
                        }
                    }
                }
                else
                {
                    foreach (var subText in subTexts)
                    {
                        subTextBuilder.AppendLine(subText);
                    }
                }

                var slide = pptPresentation.Slides.Add(newSlideNumber, layout);
                slide.Shapes[1].TextFrame.TextRange.Text = processedText;
                slide.Shapes[2].TextFrame.TextRange.Text = subTextBuilder.ToString();
                if (subTextBuilder2.Length != 0)
                {
                    slide.Shapes[3].TextFrame.TextRange.Text = subTextBuilder2.ToString();
                }
            }
        }

        private static string cleanUpText(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {

                text = text.Replace(@"\'", "'");
                text = text.Replace("\\\"", "\"");
                text = text.Trim('#');
            }

            return text;
        }
    }
}