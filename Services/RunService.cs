using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;

namespace JsonToWord.Services
{
    internal class RunService
    {
        internal Run CreateRun(WordRun wordRun)
        {
            var run = new Run();
            var runProperties = new RunProperties();

            SetHyperlink(wordRun, runProperties);
            SetBold(wordRun, runProperties);
            SetItalic(wordRun, runProperties);
            SetUnderline(wordRun, runProperties);
            SetSize(wordRun, runProperties);
            SetColor(wordRun, runProperties);

            run.AppendChild(runProperties);

            SetBreak(wordRun, run);
            SetText(wordRun, run);

            return run;
        }

        private void SetColor(WordRun wordRun, RunProperties runProperties)
        {
            if (!string.IsNullOrEmpty(wordRun.FontColor))
            {
                try
                {
                    System.Drawing.Color color = System.Drawing.Color.FromName(wordRun.FontColor);
                    string colorHex = color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
                    Color wordColor = new Color() { Val = colorHex };
                    runProperties.AppendChild(wordColor);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        private static void SetBreak(WordRun wordRun, Run run)
        {
            if (wordRun.InsertLineBreak)
                run.AppendChild(new Break());
        }

        private static void SetText(WordRun wordRun, Run run)
        {
            if (string.IsNullOrEmpty(wordRun.Text))
                return;

            var text = new Text { Text = wordRun.Text };

            if (wordRun.InsertSpace)
                text.Space = SpaceProcessingModeValues.Preserve;

            run.AppendChild(text);
        }

        private static void SetSize(WordRun wordRun, RunProperties runProperties)
        {
            if (wordRun.Size != 0)
                runProperties.FontSize = new FontSize { Val = new StringValue((wordRun.Size * 2).ToString()) };
        }

        private static void SetUnderline(WordRun wordRun, RunProperties runProperties)
        {
            if (wordRun.Underline && string.IsNullOrEmpty(wordRun.Uri))
                AddUnderline(runProperties);
        }

        private static void SetItalic(WordRun wordRun, RunProperties runProperties)
        {
            if (!wordRun.Italic)
                return;

            var italic = new Italic();
            var italicComplexScript = new ItalicComplexScript();

            runProperties.AppendChild(italic);
            runProperties.AppendChild(italicComplexScript);
        }

        private static void SetBold(WordRun wordRun, RunProperties runProperties)
        {
            if (!wordRun.Bold)
                return;

            var bold = new Bold();
            var boldComplexScript = new BoldComplexScript();

            runProperties.AppendChild(bold);
            runProperties.AppendChild(boldComplexScript);
        }

        private static void SetHyperlink(WordRun wordRun, RunProperties runProperties)
        {
            if (!string.IsNullOrEmpty(wordRun.Uri))
            {
                var runStyle = new RunStyle() { Val = "Hyperlink" };
                var color = new Color() { Val = "auto", ThemeColor = ThemeColorValues.Hyperlink };

                runProperties.AppendChild(runStyle);
                runProperties.AppendChild(color);
                AddUnderline(runProperties);
            }
            else
            {
                var runFonts = new RunFonts { Ascii = wordRun.Font, HighAnsi = wordRun.Font, ComplexScript = wordRun.Font };
                runProperties.AppendChild(runFonts);
            }
        }

        private static void AddUnderline(RunProperties runProperties)
        {
            var underline = new Underline() { Val = UnderlineValues.Single };
            runProperties.AppendChild(underline);
        }
    }
}