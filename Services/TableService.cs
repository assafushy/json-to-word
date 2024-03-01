using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;

namespace JsonToWord.Services
{
    internal class TableService
    {
        public void Insert(WordprocessingDocument document, string contentControlTitle, WordTable wordTable)
        {
            var table = CreateTable(document, wordTable);

            var contentControlService = new ContentControlService();
            var sdtBlock = contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(table);

            sdtBlock.AppendChild(sdtContentBlock);

            RemoveExtraParagraphsAfterAltChunk(document);

        }

        private Table CreateTable(WordprocessingDocument document, WordTable wordTable)
        {
            wordTable.RepeatHeaderRow = true;  

            var tableBorders = CreateTableBorders();
            var tableWidth = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };

            var tableProperties = new TableProperties();
            tableProperties.AppendChild(tableBorders);
            tableProperties.AppendChild(tableWidth);

            var isHeaderRow = true;
            var table = new Table();
            table.AppendChild(tableProperties);

            foreach (var documentRow in wordTable.Rows)
            {
                var tableRow = new TableRow { RsidTableRowProperties = "00812C40" };

                if (wordTable.RepeatHeaderRow && isHeaderRow)
                {
                    var tableHeader = new TableHeader();

                    var tableRowProperties = new TableRowProperties();
                    tableRowProperties.AppendChild(tableHeader);

                    tableRow.AppendChild(tableRowProperties);

                    isHeaderRow = false;
                }

                foreach (var cell in documentRow.Cells)
                {

                    var tableCellBorders = CreateTableCellBorders();
                    var tableCellWidth = new TableCellWidth { Width = cell.Width, Type = TableWidthUnitValues.Dxa };

                    var tableCellProperties = new TableCellProperties();
                    tableCellProperties.AppendChild(tableCellWidth);
                    tableCellProperties.AppendChild(tableCellBorders);

                    if (documentRow.MergeToOneCell)
                    {
                        var gridSpan = new GridSpan { Val = documentRow.NumberOfCellsToMerge };
                        tableCellProperties.AppendChild(gridSpan);
                    }

                    if (cell.Shading != null)
                    {
                        var cellShading = new Shading
                        {
                            Val = ShadingPatternValues.Clear,
                            Color = cell.Shading.Color,
                            Fill = cell.Shading.Fill,
                            ThemeFill = ThemeColorValues.Text2,
                            ThemeFillShade = cell.Shading.ThemeFillShade
                        };

                        tableCellProperties.AppendChild(cellShading);
                    }

                    var tableCell = new TableCell();
                    tableCell.AppendChild(tableCellProperties);
                    

                    tableCell = AppendParagraphs(tableCell, cell.Paragraphs, document);

                    tableCell = AppendAttachments(tableCell, cell.Attachments, document);


                    tableCell = AppendHtml(tableCell, cell.Html, document);

                    tableRow.AppendChild(tableCell);
                }

                table.AppendChild(tableRow);
            }

            return table;
        }

        private TableCell AppendHtml(TableCell tableCell, WordHtml html, WordprocessingDocument document)
        {

            if (html == null)
                return tableCell;

            if (string.IsNullOrEmpty(html.Html))
            {

                var paragraph = new Paragraph();
                tableCell.AppendChild(paragraph);

                return tableCell;
            }
            var styledHtml = WrapHtmlWithStyle(html.Html);

            var htmlService = new HtmlService();
            Console.WriteLine("styledHtml" + styledHtml);

            var tempHtmlFile = htmlService.CreateHtmlWordDocument(styledHtml);

            var altChunkId = "altChunkId" + Guid.NewGuid().ToString("N");
            var chunk = document.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using (var fileStream = File.Open(tempHtmlFile, FileMode.Open))
            {
                chunk.FeedData(fileStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            tableCell.AppendChild(altChunk);

            return tableCell;
        }
        private string WrapHtmlWithStyle(string originalHtml)
        {
            // This method wraps the HTML content with additional HTML tags and styles
            return $"<html style=\"font-family: Arial, sans-serif; font-size: 12pt;\"><body>{originalHtml}</body></html>";
        }
        private TableCell AppendAttachments(TableCell tableCell, List<WordAttachment> wordAttachments, WordprocessingDocument document)
        {
            if (wordAttachments == null || !wordAttachments.Any())
                return tableCell;

            var fileService = new FileService();
            var pictureService = new PictureService();

            foreach (var wordAttachment in wordAttachments)
            {
                switch (wordAttachment.Type)
                {
                    case WordObjectType.File:
                        {
                            var embeddedFile = fileService.CreateEmbeddedObject(document.MainDocumentPart, wordAttachment.Path, true);

                            var run = new Run();
                            run.AppendChild(embeddedFile);

                            var paragraph = new Paragraph();
                            paragraph.AppendChild(run);

                            tableCell.AppendChild(paragraph);
                            break;
                        }
                    case WordObjectType.Picture:
                        {
                            var drawing = pictureService.CreateDrawing(document.MainDocumentPart, wordAttachment.Path);

                            var run = new Run();
                            run.AppendChild(drawing);

                            var paragraph = new Paragraph();
                            paragraph.AppendChild(run);

                            tableCell.AppendChild(paragraph);
                            break;
                        }
                    default:
                        continue;
                }
            }

            return tableCell;
        }

        private TableCell AppendParagraphs(TableCell tableCell, List<WordParagraph> wordParagraphs, WordprocessingDocument document)
        {
            if (wordParagraphs == null || !wordParagraphs.Any())
                return tableCell;

            var paragraphService = new ParagraphService();

            foreach (var wordParagraph in wordParagraphs)
            {
                var paragraph = paragraphService.CreateParagraph(wordParagraph);

                if (wordParagraph.Runs != null && wordParagraph.Runs.Any())
                {
                    var runService = new RunService();

                    foreach (var wordRun in wordParagraph.Runs)
                    {
                        var run = runService.CreateRun(wordRun);
                        if (wordRun.Uri != null && wordRun.Uri != "")
                        {
                            try
                            {
                                var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.Uri));
                                var hyperlink = HyperlinkService.CreateHyperlink(id);
                                hyperlink.AppendChild(run);

                                paragraph.AppendChild(hyperlink);
                            }
                            catch (UriFormatException e)
                            {
                                Console.WriteLine(wordRun.Uri + " is an invalid uri \n" + e.Message);
                                paragraph.AppendChild(run);
                            }
                        }
                        else
                        {
                            paragraph.AppendChild(run);
                        }
                    }
                }

                tableCell.AppendChild(paragraph);
            }

            return tableCell;
        }

        private TableCellBorders CreateTableCellBorders()
        {
            var tableCellBorders = new TableCellBorders();
            var cellTopBorder = new TopBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellLeftBorder = new LeftBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellBottomBorder = new BottomBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var cellRightBorder = new RightBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };

            tableCellBorders.AppendChild(cellTopBorder);
            tableCellBorders.AppendChild(cellLeftBorder);
            tableCellBorders.AppendChild(cellBottomBorder);
            tableCellBorders.AppendChild(cellRightBorder);

            return tableCellBorders;
        }

        private TableBorders CreateTableBorders()
        {
            var tableBorders = new TableBorders();
            var topBorder = new TopBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var leftBorder = new LeftBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var bottomBorder = new BottomBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var rightBorder = new RightBorder { Val = BorderValues.Single, Color = "auto", Size = 4U, Space = 0U };
            var insideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 };
            var insideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single, Color = "auto", Size = 4, Space = 0 };

            tableBorders.AppendChild(topBorder);
            tableBorders.AppendChild(leftBorder);
            tableBorders.AppendChild(bottomBorder);
            tableBorders.AppendChild(rightBorder);
            tableBorders.AppendChild(insideHorizontalBorder);
            tableBorders.AppendChild(insideVerticalBorder);

            return tableBorders;
        }
         private void RemoveExtraParagraphsAfterAltChunk(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart.Document.Body;
            var altChunks = body.Descendants<AltChunk>().ToList();

            foreach (var altChunk in altChunks)
            {
                // Check for a paragraph immediately following the AltChunk
                var nextParagraph = altChunk.NextSibling<Paragraph>();
                if (nextParagraph != null)
                {
                    // Check if the paragraph is empty and if so, remove it
                    if (!nextParagraph.Descendants<Run>().Any())
                    {
                        nextParagraph.Remove();
                    }
                }

                // Check for a paragraph immediately preceding the AltChunk and remove if empty
                var prevParagraph = altChunk.PreviousSibling<Paragraph>();
                if (prevParagraph != null)
                {
                    if (!prevParagraph.Descendants<Run>().Any())
                    {
                        prevParagraph.Remove();
                    }
                }
            }
        }
    }
}