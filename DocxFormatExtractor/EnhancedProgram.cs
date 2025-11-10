using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocxFormatExtractor
{
    /// <summary>
    /// 增强版 Word 文档格式提取器
    /// 基于 Open XML SDK 官方文档和最佳实践
    /// 参考: https://learn.microsoft.com/en-us/office/open-xml/
    /// </summary>
    class EnhancedProgram
    {
        private const string DefaultDocPath = "/Users/CodeProjects/OfficeTest/test.docx";
        private const string DefaultOutputDirectory = "/Users/CodeProjects/OfficeTest";
        private const string DefaultOutputPrefix = "format_output_enhanced";
        private const string DefaultBatchInputDirectory = "/Users/CodeProjects/OfficeTest/pre_test_docx";
        private const string DefaultBatchOutputDirectory = "/Users/CodeProjects/OfficeTest/batch_output";

        private static DocumentFormatInfo formatInfo = new DocumentFormatInfo();
        private static Dictionary<string, StyleInfo> styleLookup = new Dictionary<string, StyleInfo>(StringComparer.OrdinalIgnoreCase);

        static void Main(string[] args)
        {
            if (args.Length > 0 && IsBatchArgument(args[0]))
            {
                string inputDir = args.Length > 1 ? args[1] : DefaultBatchInputDirectory;
                string outputDir = args.Length > 2 ? args[2] : DefaultBatchOutputDirectory;
                BatchDocxProcessor.Run(inputDir, outputDir);
                return;
            }

            try
            {
                var result = ProcessDocument(
                    DefaultDocPath,
                    DefaultOutputDirectory,
                    "both",
                    DefaultOutputPrefix);

                Console.WriteLine("单文件提取完成！");
                if (!string.IsNullOrEmpty(result.TextOutputPath))
                {
                    Console.WriteLine($"TXT 输出: {result.TextOutputPath}");
                }

                if (!string.IsNullOrEmpty(result.JsonOutputPath))
                {
                    Console.WriteLine($"JSON 输出: {result.JsonOutputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                Console.WriteLine($"堆栈: {ex.StackTrace}");
            }
        }

        public static ExtractionResult ProcessDocument(
            string docPath,
            string outputDirectory,
            string outputFormat = "both",
            string? outputFilePrefix = null)
        {
            if (string.IsNullOrWhiteSpace(docPath))
            {
                throw new ArgumentException("docPath 不能为空", nameof(docPath));
            }

            if (!File.Exists(docPath))
            {
                throw new FileNotFoundException($"文件不存在: {docPath}", docPath);
            }

            Directory.CreateDirectory(outputDirectory);

            formatInfo = new DocumentFormatInfo();

            Console.WriteLine("开始提取文档格式信息...");
            Console.WriteLine($"目标文件: {docPath}");
            Console.WriteLine("使用 Open XML SDK 3.x");
            Console.WriteLine();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(docPath, false))
            {
                ExtractAllInformation(doc);
            }

            string normalizedFormat = NormalizeOutputFormat(outputFormat);
            string baseName = outputFilePrefix ?? Path.GetFileNameWithoutExtension(docPath);
            string? textPath = null;
            string? jsonPath = null;

            if (normalizedFormat == "txt" || normalizedFormat == "both")
            {
                textPath = Path.Combine(outputDirectory, $"{baseName}.txt");
                OutputToText(textPath);
            }

            if (normalizedFormat == "json" || normalizedFormat == "both")
            {
                jsonPath = Path.Combine(outputDirectory, $"{baseName}.json");
                OutputToJson(jsonPath);
            }

            return new ExtractionResult(docPath, textPath, jsonPath);
        }

        private static bool IsBatchArgument(string argument)
        {
            if (string.IsNullOrWhiteSpace(argument))
            {
                return false;
            }

            var normalized = argument.TrimStart('-').ToLowerInvariant();
            return normalized == "batch" || normalized == "b";
        }

        private static string NormalizeOutputFormat(string? outputFormat)
        {
            if (string.IsNullOrWhiteSpace(outputFormat))
            {
                return "both";
            }

            string normalized = outputFormat.Trim().ToLowerInvariant();
            return normalized switch
            {
                "txt" => "txt",
                "json" => "json",
                _ => "both"
            };
        }

        static void ExtractAllInformation(WordprocessingDocument doc)
        {
            Console.WriteLine("1/10 提取文档属性...");
            ExtractDocumentProperties(doc);

            Console.WriteLine("2/10 提取样式信息...");
            ExtractStyles(doc);

            Console.WriteLine("3/10 提取段落和文本...");
            ExtractParagraphsAndRuns(doc);

            Console.WriteLine("4/10 提取表格...");
            ExtractTables(doc);

            Console.WriteLine("5/10 提取节信息...");
            ExtractSections(doc);

            Console.WriteLine("6/10 提取图片...");
            ExtractImages(doc);

            Console.WriteLine("7/10 提取页眉页脚...");
            ExtractHeadersFooters(doc);

            Console.WriteLine("8/10 提取超链接和书签...");
            ExtractHyperlinksAndBookmarks(doc);

            Console.WriteLine("9/10 提取字体和编号...");
            ExtractFontsAndNumbering(doc);

            Console.WriteLine("10/10 提取主题和批注...");
            ExtractThemesAndComments(doc);
        }

        static void ExtractDocumentProperties(WordprocessingDocument doc)
        {
            var props = new DocumentPropertiesInfo();

            // Package Properties
            var packageProps = doc.PackageProperties;
            props.Title = packageProps.Title ?? "";
            props.Subject = packageProps.Subject ?? "";
            props.Creator = packageProps.Creator ?? "";
            props.Keywords = packageProps.Keywords ?? "";
            props.Description = packageProps.Description ?? "";
            props.Category = packageProps.Category ?? "";
            props.LastModifiedBy = packageProps.LastModifiedBy ?? "";
            props.Created = packageProps.Created?.ToString() ?? "";
            props.Modified = packageProps.Modified?.ToString() ?? "";
            props.Revision = packageProps.Revision ?? "";

            // Extended Properties
            if (doc.ExtendedFilePropertiesPart?.Properties != null)
            {
                var extProps = doc.ExtendedFilePropertiesPart.Properties;
                props.Application = extProps.Application?.Text ?? "";
                props.AppVersion = extProps.ApplicationVersion?.Text ?? "";
                props.Company = extProps.Company?.Text ?? "";
                props.Manager = extProps.Manager?.Text ?? "";
                props.Pages = extProps.Pages?.Text ?? "";
                props.Words = extProps.Words?.Text ?? "";
                props.Characters = extProps.Characters?.Text ?? "";
                props.Lines = extProps.Lines?.Text ?? "";
                props.Paragraphs = extProps.Paragraphs?.Text ?? "";
            }

            // Document Settings
            if (doc.MainDocumentPart?.DocumentSettingsPart?.Settings != null)
            {
                var settings = doc.MainDocumentPart.DocumentSettingsPart.Settings;

                var zoom = settings.Elements<Zoom>().FirstOrDefault();
                if (zoom?.Percent != null)
                    props.ZoomPercent = zoom.Percent.Value.ToString();

                var defaultTabStop = settings.Elements<DefaultTabStop>().FirstOrDefault();
                if (defaultTabStop?.Val != null)
                    props.DefaultTabStop = defaultTabStop.Val.Value.ToString();

                var evenAndOddHeaders = settings.Elements<EvenAndOddHeaders>().FirstOrDefault();
                props.EvenAndOddHeaders = evenAndOddHeaders != null;
            }

            formatInfo.DocumentProperties = props;
        }

        static void ExtractStyles(WordprocessingDocument doc)
        {
            var stylesList = new List<StyleInfo>();
            var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;

            if (stylesPart?.Styles == null)
            {
                formatInfo.Styles = stylesList;
                return;
            }

            var styles = stylesPart.Styles.Elements<Style>().ToList();

            foreach (var style in styles)
            {
                var styleInfo = new StyleInfo
                {
                    StyleId = style.StyleId?.Value ?? "",
                    StyleName = style.StyleName?.Val?.Value ?? "",
                    Type = style.Type?.Value.ToString() ?? "",
                    BasedOn = style.BasedOn?.Val?.Value ?? "",
                    IsDefault = style.Default?.Value ?? false,
                    IsCustom = style.CustomStyle?.Value ?? false
                };

                // 段落属性
                if (style.StyleParagraphProperties != null)
                {
                    var paraProps = ExtractStyleParagraphProps(style.StyleParagraphProperties);
                    styleInfo.ParagraphProperties = paraProps;
                }

                // 文本属性
                if (style.StyleRunProperties != null)
                {
                    var runProps = ExtractStyleRunProps(style.StyleRunProperties);
                    styleInfo.RunProperties = runProps;
                }

                stylesList.Add(styleInfo);
            }

            formatInfo.Styles = stylesList;
            styleLookup = stylesList
                .Where(s => !string.IsNullOrEmpty(s.StyleId))
                .ToDictionary(s => s.StyleId, s => s, StringComparer.OrdinalIgnoreCase);
        }

        static void ExtractParagraphsAndRuns(WordprocessingDocument doc)
        {
            var paragraphsList = new List<ParagraphInfo>();
            var body = doc.MainDocumentPart?.Document?.Body;

            if (body == null)
            {
                formatInfo.Paragraphs = paragraphsList;
                return;
            }

            int index = 0;
            foreach (var para in body.Elements<Paragraph>())
            {
                var paraInfo = BuildParagraphInfo(para, ref index);
                paragraphsList.Add(paraInfo);
            }

            formatInfo.Paragraphs = paragraphsList;
        }

        static ParagraphInfo BuildParagraphInfo(Paragraph para, ref int index)
        {
            var paraInfo = new ParagraphInfo
            {
                Index = index++,
                Text = para.InnerText
            };

            if (para.ParagraphProperties != null)
            {
                var props = para.ParagraphProperties;
                paraInfo.StyleId = props.ParagraphStyleId?.Val?.Value ?? "";
                paraInfo.Alignment = ConvertJustificationToString(props.Justification);

                if (props.Indentation != null)
                {
                    paraInfo.LeftIndent = props.Indentation.Left?.Value ?? "";
                    paraInfo.RightIndent = props.Indentation.Right?.Value ?? "";
                    paraInfo.FirstLineIndent = props.Indentation.FirstLine?.Value ?? "";
                    paraInfo.HangingIndent = props.Indentation.Hanging?.Value ?? "";
                }

                if (props.SpacingBetweenLines != null)
                {
                    paraInfo.SpacingBefore = props.SpacingBetweenLines.Before?.Value ?? "";
                    paraInfo.SpacingAfter = props.SpacingBetweenLines.After?.Value ?? "";
                    paraInfo.LineSpacing = props.SpacingBetweenLines.Line?.Value ?? "";
                    paraInfo.LineSpacingRule = props.SpacingBetweenLines.LineRule?.Value.ToString() ?? "";
                }

                if (props.NumberingProperties != null)
                {
                    paraInfo.NumberingId = props.NumberingProperties.NumberingId?.Val?.Value.ToString() ?? "";
                    paraInfo.NumberingLevel = props.NumberingProperties.NumberingLevelReference?.Val?.Value.ToString() ?? "";
                }

                if (props.ParagraphBorders != null)
                {
                    paraInfo.HasBorders = true;
                }

                if (props.Shading != null)
                {
                    paraInfo.ShadingFill = props.Shading.Fill?.Value ?? "";
                    paraInfo.ShadingColor = props.Shading.Color?.Value ?? "";
                }
            }

            var runs = new List<RunInfo>();
            foreach (var run in para.Elements<Run>())
            {
                var runInfo = new RunInfo
                {
                    Text = run.InnerText
                };

                if (run.RunProperties != null)
                {
                    var rProps = run.RunProperties;

                    if (rProps.RunFonts != null)
                    {
                        runInfo.FontNameAscii = rProps.RunFonts.Ascii?.Value ?? "";
                        runInfo.FontNameEastAsia = rProps.RunFonts.EastAsia?.Value ?? "";
                        runInfo.FontNameComplexScript = rProps.RunFonts.ComplexScript?.Value ?? "";
                    }

                    runInfo.FontSize = rProps.FontSize?.Val?.Value ?? "";
                    runInfo.Bold = rProps.Bold != null;
                    runInfo.Italic = rProps.Italic != null;
                    runInfo.Underline = rProps.Underline?.Val?.Value.ToString() ?? "";
                    runInfo.Strike = rProps.Strike != null;
                    runInfo.Color = rProps.Color?.Val?.Value ?? "";
                    runInfo.Highlight = rProps.Highlight?.Val?.Value.ToString() ?? "";
                    runInfo.VerticalAlignment = rProps.VerticalTextAlignment?.Val?.Value.ToString() ?? "";
                }

                runs.Add(runInfo);
            }
            paraInfo.Runs = runs;

            ApplyParagraphStyleFallbacks(paraInfo);

            return paraInfo;
        }

        static void ExtractTables(WordprocessingDocument doc)
        {
            var tablesList = new List<TableInfo>();
            var body = doc.MainDocumentPart?.Document?.Body;

            if (body == null)
            {
                formatInfo.Tables = tablesList;
                return;
            }

            int index = 0;
            foreach (var table in body.Elements<Table>())
            {
                var tableInfo = new TableInfo
                {
                    Index = index++
                };

                var tableProps = table.GetFirstChild<TableProperties>();
                if (tableProps != null)
                {
                    tableInfo.StyleId = tableProps.TableStyle?.Val?.Value ?? "";
                    tableInfo.Width = tableProps.TableWidth?.Width?.Value ?? "";
                    tableInfo.WidthType = tableProps.TableWidth?.Type?.Value.ToString() ?? "";
                    tableInfo.Alignment = tableProps.TableJustification?.Val?.Value.ToString() ?? "";
                    tableInfo.HasBorders = tableProps.TableBorders != null;
                }

                // 提取行和单元格
                var rows = new List<TableRowInfo>();
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowInfo = new TableRowInfo();

                    var rowProps = row.GetFirstChild<TableRowProperties>();
                    if (rowProps != null)
                    {
                        var height = rowProps.GetFirstChild<TableRowHeight>();
                        rowInfo.Height = height?.Val.HasValue == true ? height.Val.Value.ToString() : "";
                        rowInfo.IsHeader = rowProps.GetFirstChild<TableHeader>() != null;
                    }

                    var cells = new List<TableCellInfo>();
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        var cellInfo = new TableCellInfo
                        {
                            Text = cell.InnerText
                        };

                        var cellProps = cell.GetFirstChild<TableCellProperties>();
                        if (cellProps != null)
                        {
                            cellInfo.Width = cellProps.TableCellWidth?.Width?.Value ?? "";
                            cellInfo.VerticalAlignment = cellProps.TableCellVerticalAlignment?.Val?.Value.ToString() ?? "";

                            if (cellProps.Shading != null)
                            {
                                cellInfo.BackgroundColor = cellProps.Shading.Fill?.Value ?? "";
                            }

                            cellInfo.VerticalMerge = cellProps.VerticalMerge?.Val?.Value.ToString() ?? "";
                            cellInfo.HorizontalMerge = cellProps.HorizontalMerge?.Val?.Value.ToString() ?? "";
                        }

                        cells.Add(cellInfo);
                    }
                    rowInfo.Cells = cells;
                    rows.Add(rowInfo);
                }
                tableInfo.Rows = rows;

                tablesList.Add(tableInfo);
            }

            formatInfo.Tables = tablesList;
        }

        static void ExtractSections(WordprocessingDocument doc)
        {
            var sectionsList = new List<SectionInfo>();
            var body = doc.MainDocumentPart?.Document?.Body;

            if (body == null)
            {
                formatInfo.Sections = sectionsList;
                return;
            }

            // 收集所有节属性
            var allSections = new List<SectionProperties>();

            // 检查body直接子元素的SectionProperties
            allSections.AddRange(body.Elements<SectionProperties>());

            // 检查最后一个段落的SectionProperties（这是Word存储最后一节的地方）
            var lastPara = body.Elements<Paragraph>().LastOrDefault();
            if (lastPara?.ParagraphProperties?.SectionProperties != null)
            {
                allSections.Add(lastPara.ParagraphProperties.SectionProperties);
            }

            int index = 0;
            foreach (var section in allSections)
            {
                var sectionInfo = new SectionInfo
                {
                    Index = index++
                };

                var pageSize = section.GetFirstChild<PageSize>();
                if (pageSize != null)
                {
                    sectionInfo.PageWidth = pageSize.Width.HasValue ? pageSize.Width.Value.ToString() : "";
                    sectionInfo.PageHeight = pageSize.Height.HasValue ? pageSize.Height.Value.ToString() : "";
                    sectionInfo.Orientation = pageSize.Orient?.Value.ToString() ?? "Portrait";
                }

                var pageMargin = section.GetFirstChild<PageMargin>();
                if (pageMargin != null)
                {
                    sectionInfo.MarginTop = pageMargin.Top.HasValue ? pageMargin.Top.Value.ToString() : "";
                    sectionInfo.MarginBottom = pageMargin.Bottom.HasValue ? pageMargin.Bottom.Value.ToString() : "";
                    sectionInfo.MarginLeft = pageMargin.Left.HasValue ? pageMargin.Left.Value.ToString() : "";
                    sectionInfo.MarginRight = pageMargin.Right.HasValue ? pageMargin.Right.Value.ToString() : "";
                    sectionInfo.MarginHeader = pageMargin.Header.HasValue ? pageMargin.Header.Value.ToString() : "";
                    sectionInfo.MarginFooter = pageMargin.Footer.HasValue ? pageMargin.Footer.Value.ToString() : "";
                    sectionInfo.MarginGutter = pageMargin.Gutter.HasValue ? pageMargin.Gutter.Value.ToString() : "";
                }

                var columns = section.GetFirstChild<Columns>();
                if (columns != null)
                {
                    sectionInfo.ColumnCount = columns.ColumnCount?.Value.ToString() ?? "1";
                    sectionInfo.ColumnSpacing = columns.Space?.Value ?? "";
                }

                var sectionType = section.GetFirstChild<SectionType>();
                if (sectionType != null)
                {
                    sectionInfo.SectionType = sectionType.Val?.Value.ToString() ?? "";
                }

                sectionsList.Add(sectionInfo);
            }

            formatInfo.Sections = sectionsList;
        }

        static void ExtractImages(WordprocessingDocument doc)
        {
            var imagesList = new List<ImageInfo>();

            if (doc.MainDocumentPart == null)
            {
                formatInfo.Images = imagesList;
                return;
            }

            int index = 0;

            // 方法1: 直接从 ImageParts 提取
            foreach (var imagePart in doc.MainDocumentPart.ImageParts)
            {
                var imageInfo = new ImageInfo
                {
                    Index = index++,
                    ContentType = imagePart.ContentType,
                    RelationshipId = doc.MainDocumentPart.GetIdOfPart(imagePart)
                };

                // 获取图片大小
                using (var stream = imagePart.GetStream())
                {
                    imageInfo.SizeBytes = stream.Length;
                }

                imagesList.Add(imageInfo);
            }

            // 方法2: 从文档中查找Drawing元素获取更多信息
            var body = doc.MainDocumentPart.Document?.Body;
            if (body != null)
            {
                var drawings = body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().ToList();

                for (int i = 0; i < Math.Min(drawings.Count, imagesList.Count); i++)
                {
                    var drawing = drawings[i];
                    var extent = drawing.Extent;
                    if (extent != null)
                    {
                        imagesList[i].Width = extent.Cx.HasValue ? extent.Cx.Value.ToString() : "";
                        imagesList[i].Height = extent.Cy.HasValue ? extent.Cy.Value.ToString() : "";
                    }

                    var docProperties = drawing.DocProperties;
                    if (docProperties != null)
                    {
                        imagesList[i].Name = docProperties.Name?.Value ?? "";
                        imagesList[i].Description = docProperties.Description?.Value ?? "";
                    }
                }
            }

            formatInfo.Images = imagesList;
        }

        static void ExtractHeadersFooters(WordprocessingDocument doc)
        {
            var headersList = new List<HeaderFooterInfo>();
            var footersList = new List<HeaderFooterInfo>();

            if (doc.MainDocumentPart == null)
            {
                formatInfo.Headers = headersList;
                formatInfo.Footers = footersList;
                return;
            }

            // 提取页眉
            int headerIndex = 0;
            foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
            {
                var headerInfo = new HeaderFooterInfo
                {
                    Index = headerIndex++,
                    Text = headerPart.Header?.InnerText ?? "",
                    RelationshipId = doc.MainDocumentPart.GetIdOfPart(headerPart)
                };

                if (headerPart.Header != null)
                {
                    int paragraphIndex = 0;
                    foreach (var paragraph in headerPart.Header.Elements<Paragraph>())
                    {
                        var paragraphInfo = BuildParagraphInfo(paragraph, ref paragraphIndex);
                        headerInfo.Paragraphs.Add(paragraphInfo);
                    }
                }

                headersList.Add(headerInfo);
            }

            // 提取页脚
            int footerIndex = 0;
            foreach (var footerPart in doc.MainDocumentPart.FooterParts)
            {
                var footerInfo = new HeaderFooterInfo
                {
                    Index = footerIndex++,
                    Text = footerPart.Footer?.InnerText ?? "",
                    RelationshipId = doc.MainDocumentPart.GetIdOfPart(footerPart)
                };

                if (footerPart.Footer != null)
                {
                    int paragraphIndex = 0;
                    foreach (var paragraph in footerPart.Footer.Elements<Paragraph>())
                    {
                        var paragraphInfo = BuildParagraphInfo(paragraph, ref paragraphIndex);
                        footerInfo.Paragraphs.Add(paragraphInfo);
                    }
                }

                footersList.Add(footerInfo);
            }

            formatInfo.Headers = headersList;
            formatInfo.Footers = footersList;
        }

        static void ExtractHyperlinksAndBookmarks(WordprocessingDocument doc)
        {
            var hyperlinksList = new List<HyperlinkInfo>();
            var bookmarksList = new List<BookmarkInfo>();

            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                formatInfo.Hyperlinks = hyperlinksList;
                formatInfo.Bookmarks = bookmarksList;
                return;
            }

            // 提取超链接
            int hyperlinkIndex = 0;
            foreach (var hyperlink in body.Descendants<Hyperlink>())
            {
                var hyperlinkInfo = new HyperlinkInfo
                {
                    Index = hyperlinkIndex++,
                    Text = hyperlink.InnerText,
                    Anchor = hyperlink.Anchor?.Value ?? ""
                };

                // 获取外部链接URL
                if (hyperlink.Id != null)
                {
                    var relationship = doc.MainDocumentPart.HyperlinkRelationships
                        .FirstOrDefault(r => r.Id == hyperlink.Id.Value);

                    if (relationship != null)
                    {
                        hyperlinkInfo.Url = relationship.Uri?.ToString() ?? "";
                        hyperlinkInfo.IsExternal = relationship.IsExternal;
                    }
                }

                hyperlinksList.Add(hyperlinkInfo);
            }

            // 提取书签
            int bookmarkIndex = 0;
            foreach (var bookmarkStart in body.Descendants<BookmarkStart>())
            {
                var bookmarkInfo = new BookmarkInfo
                {
                    Index = bookmarkIndex++,
                    Id = bookmarkStart.Id?.Value ?? "",
                    Name = bookmarkStart.Name?.Value ?? ""
                };

                bookmarksList.Add(bookmarkInfo);
            }

            formatInfo.Hyperlinks = hyperlinksList;
            formatInfo.Bookmarks = bookmarksList;
        }

        static void ExtractFontsAndNumbering(WordprocessingDocument doc)
        {
            // 字体表
            var fontsList = new List<FontInfo>();
            var fontTablePart = doc.MainDocumentPart?.FontTablePart;

            if (fontTablePart?.Fonts != null)
            {
                foreach (var font in fontTablePart.Fonts.Elements<DocumentFormat.OpenXml.Wordprocessing.Font>())
                {
                    var fontInfo = new FontInfo
                    {
                        Name = font.Name?.Value ?? ""
                    };

                    var fontFamily = font.GetFirstChild<FontFamily>();
                    if (fontFamily != null)
                    {
                        fontInfo.Family = fontFamily.Val?.Value.ToString() ?? "";
                    }

                    var pitch = font.GetFirstChild<Pitch>();
                    if (pitch != null)
                    {
                        fontInfo.Pitch = pitch.Val?.Value.ToString() ?? "";
                    }

                    fontsList.Add(fontInfo);
                }
            }
            formatInfo.Fonts = fontsList;

            // 编号定义
            var numberingList = new List<NumberingInfo>();
            var numberingPart = doc.MainDocumentPart?.NumberingDefinitionsPart;

            if (numberingPart?.Numbering != null)
            {
                foreach (var abstractNum in numberingPart.Numbering.Elements<AbstractNum>())
                {
                    var numInfo = new NumberingInfo
                    {
                        AbstractNumId = abstractNum.AbstractNumberId?.Value.ToString() ?? "",
                        LevelCount = abstractNum.Elements<Level>().Count()
                    };

                    var levels = new List<NumberingLevelInfo>();
                    foreach (var level in abstractNum.Elements<Level>())
                    {
                        var levelInfo = new NumberingLevelInfo
                        {
                            LevelIndex = level.LevelIndex?.Value.ToString() ?? "",
                            NumberFormat = level.NumberingFormat?.Val?.Value.ToString() ?? "",
                            LevelText = level.LevelText?.Val?.Value ?? "",
                            StartValue = level.StartNumberingValue?.Val?.Value.ToString() ?? ""
                        };

                        levels.Add(levelInfo);
                    }
                    numInfo.Levels = levels;

                    numberingList.Add(numInfo);
                }
            }
            formatInfo.Numbering = numberingList;
        }

        static void ExtractThemesAndComments(WordprocessingDocument doc)
        {
            // 主题
            var themePart = doc.MainDocumentPart?.ThemePart;
            if (themePart?.Theme != null)
            {
                var theme = themePart.Theme;
                formatInfo.ThemeName = theme.Name?.Value ?? "";
            }

            // 批注
            var commentsList = new List<CommentInfo>();
            var commentsPart = doc.MainDocumentPart?.WordprocessingCommentsPart;

            if (commentsPart?.Comments != null)
            {
                int index = 0;
                foreach (var comment in commentsPart.Comments.Elements<Comment>())
                {
                    var commentInfo = new CommentInfo
                    {
                        Index = index++,
                        Id = comment.Id?.Value ?? "",
                        Author = comment.Author?.Value ?? "",
                        Date = comment.Date?.Value.ToString() ?? "",
                        Text = comment.InnerText
                    };

                    commentsList.Add(commentInfo);
                }
            }
            formatInfo.Comments = commentsList;
        }

        static ParagraphPropertiesInfo ExtractStyleParagraphProps(StyleParagraphProperties props)
        {
            var info = new ParagraphPropertiesInfo
            {
                Alignment = ConvertJustificationToString(props.Justification)
            };

            if (props.Indentation != null)
            {
                info.LeftIndent = props.Indentation.Left?.Value ?? "";
                info.RightIndent = props.Indentation.Right?.Value ?? "";
                info.FirstLineIndent = props.Indentation.FirstLine?.Value ?? "";
            }

            if (props.SpacingBetweenLines != null)
            {
                info.SpacingBefore = props.SpacingBetweenLines.Before?.Value ?? "";
                info.SpacingAfter = props.SpacingBetweenLines.After?.Value ?? "";
                info.LineSpacing = props.SpacingBetweenLines.Line?.Value ?? "";
            }

            return info;
        }

        static RunPropertiesInfo ExtractStyleRunProps(StyleRunProperties props)
        {
            var info = new RunPropertiesInfo();

            if (props.RunFonts != null)
            {
                info.FontNameAscii = props.RunFonts.Ascii?.Value ?? "";
                info.FontNameEastAsia = props.RunFonts.EastAsia?.Value ?? "";
            }

            info.FontSize = props.FontSize?.Val?.Value ?? "";
            info.Bold = props.Bold != null;
            info.Italic = props.Italic != null;
            info.Color = props.Color?.Val?.Value ?? "";

            return info;
        }

        static string ConvertJustificationToString(Justification? justification)
        {
            if (justification?.Val == null)
            {
                return "";
            }

            var innerText = justification.Val.InnerText;
            if (!string.IsNullOrWhiteSpace(innerText))
            {
                return innerText.Trim();
            }

            try
            {
                return Enum.GetName(typeof(JustificationValues), justification.Val.Value) ?? "";
            }
            catch
            {
                return "";
            }
        }

        static void ApplyParagraphStyleFallbacks(ParagraphInfo paraInfo)
        {
            if (styleLookup.Count == 0)
            {
                return;
            }

            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? currentStyleId = string.IsNullOrEmpty(paraInfo.StyleId) ? "Normal" : paraInfo.StyleId;

            while (!string.IsNullOrEmpty(currentStyleId))
            {
                if (!visited.Add(currentStyleId))
                {
                    break;
                }

                if (!styleLookup.TryGetValue(currentStyleId, out var style))
                {
                    break;
                }

                if (style.ParagraphProperties != null)
                {
                    var props = style.ParagraphProperties;
                    paraInfo.Alignment = UseFallback(paraInfo.Alignment, props.Alignment);
                    paraInfo.LeftIndent = UseFallback(paraInfo.LeftIndent, props.LeftIndent);
                    paraInfo.RightIndent = UseFallback(paraInfo.RightIndent, props.RightIndent);
                    paraInfo.FirstLineIndent = UseFallback(paraInfo.FirstLineIndent, props.FirstLineIndent);
                    paraInfo.SpacingBefore = UseFallback(paraInfo.SpacingBefore, props.SpacingBefore);
                    paraInfo.SpacingAfter = UseFallback(paraInfo.SpacingAfter, props.SpacingAfter);
                    paraInfo.LineSpacing = UseFallback(paraInfo.LineSpacing, props.LineSpacing);
                }

                currentStyleId = style.BasedOn;
            }
        }

        static string UseFallback(string current, string? fallback)
        {
            return string.IsNullOrEmpty(current) ? (fallback ?? "") : current;
        }

        static void OutputToText(string filePath)
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine("=" + new string('=', 100));
                writer.WriteLine("Word 文档完整格式信息提取报告");
                writer.WriteLine("提取工具: Open XML SDK 3.x 增强版");
                writer.WriteLine("生成时间: " + DateTime.Now);
                writer.WriteLine("=" + new string('=', 100));
                writer.WriteLine();

                // 文档属性
                writer.WriteLine("【1. 文档属性】");
                writer.WriteLine(new string('-', 100));
                var props = formatInfo.DocumentProperties;
                writer.WriteLine($"标题: {props.Title}");
                writer.WriteLine($"主题: {props.Subject}");
                writer.WriteLine($"创建者: {props.Creator}");
                writer.WriteLine($"最后修改者: {props.LastModifiedBy}");
                writer.WriteLine($"创建时间: {props.Created}");
                writer.WriteLine($"修改时间: {props.Modified}");
                writer.WriteLine($"版本: {props.Revision}");
                writer.WriteLine($"应用程序: {props.Application}");
                writer.WriteLine($"公司: {props.Company}");
                writer.WriteLine($"页数: {props.Pages}");
                writer.WriteLine($"字数: {props.Words}");
                writer.WriteLine($"字符数: {props.Characters}");
                writer.WriteLine($"段落数: {props.Paragraphs}");
                writer.WriteLine();

                // 样式
                writer.WriteLine($"【2. 样式】 总数: {formatInfo.Styles.Count}");
                writer.WriteLine(new string('-', 100));
                foreach (var style in formatInfo.Styles.Take(20)) // 仅显示前20个
                {
                    writer.WriteLine($"样式: {style.StyleName} (ID: {style.StyleId})");
                    writer.WriteLine($"  类型: {style.Type}, 基于: {style.BasedOn}");
                    if (style.RunProperties != null && !string.IsNullOrEmpty(style.RunProperties.FontSize))
                    {
                        writer.WriteLine($"  字号: {style.RunProperties.FontSize}, 粗体: {style.RunProperties.Bold}, 颜色: {style.RunProperties.Color}");
                    }
                }
                writer.WriteLine($"... (共 {formatInfo.Styles.Count} 个样式)");
                writer.WriteLine();

                // 段落统计
                writer.WriteLine($"【3. 段落】 总数: {formatInfo.Paragraphs.Count}");
                writer.WriteLine(new string('-', 100));
                writer.WriteLine($"前10个段落预览:");
                foreach (var para in formatInfo.Paragraphs.Take(10))
                {
                    writer.WriteLine($"段落 #{para.Index}: {para.Text.Substring(0, Math.Min(50, para.Text.Length))}...");
                    writer.WriteLine($"  样式: {para.StyleId}, 对齐: {para.Alignment}");
                }
                writer.WriteLine();

                // 表格
                writer.WriteLine($"【4. 表格】 总数: {formatInfo.Tables.Count}");
                writer.WriteLine(new string('-', 100));
                foreach (var table in formatInfo.Tables)
                {
                    writer.WriteLine($"表格 #{table.Index}: {table.Rows.Count}行");
                    writer.WriteLine($"  样式: {table.StyleId}, 对齐: {table.Alignment}");
                }
                writer.WriteLine();

                // 图片
                writer.WriteLine($"【5. 图片】 总数: {formatInfo.Images.Count}");
                writer.WriteLine(new string('-', 100));
                foreach (var image in formatInfo.Images)
                {
                    writer.WriteLine($"图片 #{image.Index}:");
                    writer.WriteLine($"  类型: {image.ContentType}");
                    writer.WriteLine($"  大小: {image.SizeBytes} 字节");
                    writer.WriteLine($"  名称: {image.Name}");
                    writer.WriteLine($"  尺寸: {image.Width} x {image.Height}");
                }
                writer.WriteLine();

                // 节
                writer.WriteLine($"【6. 节】 总数: {formatInfo.Sections.Count}");
                writer.WriteLine(new string('-', 100));
                foreach (var section in formatInfo.Sections)
                {
                    writer.WriteLine($"节 #{section.Index}:");
                    writer.WriteLine($"  页面: {section.PageWidth} x {section.PageHeight}, 方向: {section.Orientation}");
                    writer.WriteLine($"  边距: 上{section.MarginTop} 下{section.MarginBottom} 左{section.MarginLeft} 右{section.MarginRight}");
                }
                writer.WriteLine();

                // 超链接和书签
                writer.WriteLine($"【7. 超链接】 总数: {formatInfo.Hyperlinks.Count}");
                foreach (var link in formatInfo.Hyperlinks.Take(10))
                {
                    writer.WriteLine($"  {link.Text} -> {link.Url}");
                }
                writer.WriteLine();

                writer.WriteLine($"【8. 书签】 总数: {formatInfo.Bookmarks.Count}");
                foreach (var bookmark in formatInfo.Bookmarks.Take(10))
                {
                    writer.WriteLine($"  {bookmark.Name} (ID: {bookmark.Id})");
                }
                writer.WriteLine();

                // 页眉页脚
                writer.WriteLine($"【9. 页眉】 总数: {formatInfo.Headers.Count}");
                foreach (var header in formatInfo.Headers)
                {
                    writer.WriteLine($"  页眉 #{header.Index}: {header.Text}");
                }
                writer.WriteLine();

                writer.WriteLine($"【10. 页脚】 总数: {formatInfo.Footers.Count}");
                foreach (var footer in formatInfo.Footers)
                {
                    writer.WriteLine($"  页脚 #{footer.Index}: {footer.Text}");
                }
                writer.WriteLine();

                // 批注
                writer.WriteLine($"【11. 批注】 总数: {formatInfo.Comments.Count}");
                foreach (var comment in formatInfo.Comments)
                {
                    writer.WriteLine($"  批注 by {comment.Author} ({comment.Date}): {comment.Text}");
                }
                writer.WriteLine();

                // 主题
                writer.WriteLine($"【12. 主题】 {formatInfo.ThemeName}");
                writer.WriteLine();

                writer.WriteLine("=" + new string('=', 100));
                writer.WriteLine("报告结束");
                writer.WriteLine("=" + new string('=', 100));
            }

            Console.WriteLine($"文本报告已保存到: {filePath}");
        }

        static void OutputToJson(string filePath)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string json = JsonSerializer.Serialize(formatInfo, options);
            File.WriteAllText(filePath, json);

            Console.WriteLine($"JSON报告已保存到: {filePath}");
        }
    }

    public record ExtractionResult(string DocumentPath, string? TextOutputPath, string? JsonOutputPath);

    #region 数据模型类

    public class DocumentFormatInfo
    {
        public DocumentPropertiesInfo DocumentProperties { get; set; } = new DocumentPropertiesInfo();
        public List<StyleInfo> Styles { get; set; } = new List<StyleInfo>();
        public List<ParagraphInfo> Paragraphs { get; set; } = new List<ParagraphInfo>();
        public List<TableInfo> Tables { get; set; } = new List<TableInfo>();
        public List<SectionInfo> Sections { get; set; } = new List<SectionInfo>();
        public List<ImageInfo> Images { get; set; } = new List<ImageInfo>();
        public List<HeaderFooterInfo> Headers { get; set; } = new List<HeaderFooterInfo>();
        public List<HeaderFooterInfo> Footers { get; set; } = new List<HeaderFooterInfo>();
        public List<HyperlinkInfo> Hyperlinks { get; set; } = new List<HyperlinkInfo>();
        public List<BookmarkInfo> Bookmarks { get; set; } = new List<BookmarkInfo>();
        public List<FontInfo> Fonts { get; set; } = new List<FontInfo>();
        public List<NumberingInfo> Numbering { get; set; } = new List<NumberingInfo>();
        public List<CommentInfo> Comments { get; set; } = new List<CommentInfo>();
        public string ThemeName { get; set; } = "";
    }

    public class DocumentPropertiesInfo
    {
        public string Title { get; set; } = "";
        public string Subject { get; set; } = "";
        public string Creator { get; set; } = "";
        public string Keywords { get; set; } = "";
        public string Description { get; set; } = "";
        public string Category { get; set; } = "";
        public string LastModifiedBy { get; set; } = "";
        public string Created { get; set; } = "";
        public string Modified { get; set; } = "";
        public string Revision { get; set; } = "";
        public string Application { get; set; } = "";
        public string AppVersion { get; set; } = "";
        public string Company { get; set; } = "";
        public string Manager { get; set; } = "";
        public string Pages { get; set; } = "";
        public string Words { get; set; } = "";
        public string Characters { get; set; } = "";
        public string Lines { get; set; } = "";
        public string Paragraphs { get; set; } = "";
        public string ZoomPercent { get; set; } = "";
        public string DefaultTabStop { get; set; } = "";
        public bool EvenAndOddHeaders { get; set; } = false;
    }

    public class StyleInfo
    {
        public string StyleId { get; set; } = "";
        public string StyleName { get; set; } = "";
        public string Type { get; set; } = "";
        public string BasedOn { get; set; } = "";
        public bool IsDefault { get; set; }
        public bool IsCustom { get; set; }
        public ParagraphPropertiesInfo? ParagraphProperties { get; set; }
        public RunPropertiesInfo? RunProperties { get; set; }
    }

    public class ParagraphPropertiesInfo
    {
        public string Alignment { get; set; } = "";
        public string LeftIndent { get; set; } = "";
        public string RightIndent { get; set; } = "";
        public string FirstLineIndent { get; set; } = "";
        public string SpacingBefore { get; set; } = "";
        public string SpacingAfter { get; set; } = "";
        public string LineSpacing { get; set; } = "";
    }

    public class RunPropertiesInfo
    {
        public string FontNameAscii { get; set; } = "";
        public string FontNameEastAsia { get; set; } = "";
        public string FontSize { get; set; } = "";
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string Color { get; set; } = "";
    }

    public class ParagraphInfo
    {
        public int Index { get; set; }
        public string Text { get; set; } = "";
        public string StyleId { get; set; } = "";
        public string Alignment { get; set; } = "";
        public string LeftIndent { get; set; } = "";
        public string RightIndent { get; set; } = "";
        public string FirstLineIndent { get; set; } = "";
        public string HangingIndent { get; set; } = "";
        public string SpacingBefore { get; set; } = "";
        public string SpacingAfter { get; set; } = "";
        public string LineSpacing { get; set; } = "";
        public string LineSpacingRule { get; set; } = "";
        public string NumberingId { get; set; } = "";
        public string NumberingLevel { get; set; } = "";
        public bool HasBorders { get; set; }
        public string ShadingFill { get; set; } = "";
        public string ShadingColor { get; set; } = "";
        public List<RunInfo> Runs { get; set; } = new List<RunInfo>();
    }

    public class RunInfo
    {
        public string Text { get; set; } = "";
        public string FontNameAscii { get; set; } = "";
        public string FontNameEastAsia { get; set; } = "";
        public string FontNameComplexScript { get; set; } = "";
        public string FontSize { get; set; } = "";
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string Underline { get; set; } = "";
        public bool Strike { get; set; }
        public string Color { get; set; } = "";
        public string Highlight { get; set; } = "";
        public string VerticalAlignment { get; set; } = "";
    }

    public class TableInfo
    {
        public int Index { get; set; }
        public string StyleId { get; set; } = "";
        public string Width { get; set; } = "";
        public string WidthType { get; set; } = "";
        public string Alignment { get; set; } = "";
        public bool HasBorders { get; set; }
        public List<TableRowInfo> Rows { get; set; } = new List<TableRowInfo>();
    }

    public class TableRowInfo
    {
        public string Height { get; set; } = "";
        public bool IsHeader { get; set; }
        public List<TableCellInfo> Cells { get; set; } = new List<TableCellInfo>();
    }

    public class TableCellInfo
    {
        public string Text { get; set; } = "";
        public string Width { get; set; } = "";
        public string VerticalAlignment { get; set; } = "";
        public string BackgroundColor { get; set; } = "";
        public string VerticalMerge { get; set; } = "";
        public string HorizontalMerge { get; set; } = "";
    }

    public class SectionInfo
    {
        public int Index { get; set; }
        public string PageWidth { get; set; } = "";
        public string PageHeight { get; set; } = "";
        public string Orientation { get; set; } = "";
        public string MarginTop { get; set; } = "";
        public string MarginBottom { get; set; } = "";
        public string MarginLeft { get; set; } = "";
        public string MarginRight { get; set; } = "";
        public string MarginHeader { get; set; } = "";
        public string MarginFooter { get; set; } = "";
        public string MarginGutter { get; set; } = "";
        public string ColumnCount { get; set; } = "";
        public string ColumnSpacing { get; set; } = "";
        public string SectionType { get; set; } = "";
    }

    public class ImageInfo
    {
        public int Index { get; set; }
        public string ContentType { get; set; } = "";
        public string RelationshipId { get; set; } = "";
        public long SizeBytes { get; set; }
        public string Width { get; set; } = "";
        public string Height { get; set; } = "";
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
    }

    public class HeaderFooterInfo
    {
        public int Index { get; set; }
        public string Text { get; set; } = "";
        public string RelationshipId { get; set; } = "";
        public List<ParagraphInfo> Paragraphs { get; set; } = new List<ParagraphInfo>();
    }

    public class HyperlinkInfo
    {
        public int Index { get; set; }
        public string Text { get; set; } = "";
        public string Url { get; set; } = "";
        public string Anchor { get; set; } = "";
        public bool IsExternal { get; set; }
    }

    public class BookmarkInfo
    {
        public int Index { get; set; }
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
    }

    public class FontInfo
    {
        public string Name { get; set; } = "";
        public string Family { get; set; } = "";
        public string Pitch { get; set; } = "";
    }

    public class NumberingInfo
    {
        public string AbstractNumId { get; set; } = "";
        public int LevelCount { get; set; }
        public List<NumberingLevelInfo> Levels { get; set; } = new List<NumberingLevelInfo>();
    }

    public class NumberingLevelInfo
    {
        public string LevelIndex { get; set; } = "";
        public string NumberFormat { get; set; } = "";
        public string LevelText { get; set; } = "";
        public string StartValue { get; set; } = "";
    }

    public class CommentInfo
    {
        public int Index { get; set; }
        public string Id { get; set; } = "";
        public string Author { get; set; } = "";
        public string Date { get; set; } = "";
        public string Text { get; set; } = "";
    }

    #endregion
}
