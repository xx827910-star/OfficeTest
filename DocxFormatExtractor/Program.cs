using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        string docPath = "/Users/CodeProjects/OfficeTest/test.docx";

        if (!File.Exists(docPath))
        {
            Console.WriteLine($"文件不存在: {docPath}");
            return;
        }

        Console.WriteLine("=" + new string('=', 100));
        Console.WriteLine($"提取文档格式信息: {docPath}");
        Console.WriteLine("=" + new string('=', 100));
        Console.WriteLine();

        using (WordprocessingDocument doc = WordprocessingDocument.Open(docPath, false))
        {
            // 1. 文档属性
            ExtractDocumentProperties(doc);

            // 2. 样式信息
            ExtractStyles(doc);

            // 3. 主文档内容格式
            ExtractMainDocumentFormat(doc);

            // 4. 节信息
            ExtractSections(doc);

            // 5. 字体信息
            ExtractFonts(doc);

            // 6. 编号和列表
            ExtractNumbering(doc);
        }

        Console.WriteLine();
        Console.WriteLine("=" + new string('=', 100));
        Console.WriteLine("提取完成");
        Console.WriteLine("=" + new string('=', 100));
    }

    static void ExtractDocumentProperties(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【1. 文档属性】");
        Console.WriteLine(new string('-', 100));

        var docProps = doc.PackageProperties;
        Console.WriteLine($"标题: {docProps.Title ?? "无"}");
        Console.WriteLine($"主题: {docProps.Subject ?? "无"}");
        Console.WriteLine($"创建者: {docProps.Creator ?? "无"}");
        Console.WriteLine($"最后修改者: {docProps.LastModifiedBy ?? "无"}");
        Console.WriteLine($"创建时间: {docProps.Created}");
        Console.WriteLine($"修改时间: {docProps.Modified}");
        Console.WriteLine($"关键词: {docProps.Keywords ?? "无"}");
        Console.WriteLine($"描述: {docProps.Description ?? "无"}");
        Console.WriteLine($"类别: {docProps.Category ?? "无"}");

        var mainPart = doc.MainDocumentPart;
        if (mainPart?.DocumentSettingsPart?.Settings != null)
        {
            var settings = mainPart.DocumentSettingsPart.Settings;
            Console.WriteLine("\n文档设置:");

            var zoom = settings.Elements<Zoom>().FirstOrDefault();
            if (zoom != null && zoom.Percent != null)
            {
                Console.WriteLine($"  缩放比例: {zoom.Percent.Value}%");
            }

            var defaultTabStop = settings.Elements<DefaultTabStop>().FirstOrDefault();
            if (defaultTabStop != null)
            {
                Console.WriteLine($"  默认制表位: {defaultTabStop.Val?.Value}");
            }
        }
    }

    static void ExtractStyles(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【2. 样式信息】");
        Console.WriteLine(new string('-', 100));

        var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null)
        {
            Console.WriteLine("无样式定义");
            return;
        }

        var styles = stylesPart.Styles.Elements<Style>().ToList();
        Console.WriteLine($"样式总数: {styles.Count}");
        Console.WriteLine();

        foreach (var style in styles)
        {
            Console.WriteLine($"样式ID: {style.StyleId?.Value ?? "无ID"}");
            Console.WriteLine($"  类型: {style.Type?.Value}");
            Console.WriteLine($"  名称: {style.StyleName?.Val?.Value ?? "无名称"}");
            Console.WriteLine($"  基于: {style.BasedOn?.Val?.Value ?? "无"}");
            Console.WriteLine($"  默认: {style.Default?.Value ?? false}");
            Console.WriteLine($"  自定义: {style.CustomStyle?.Value ?? false}");

            // 段落属性
            if (style.StyleParagraphProperties != null)
            {
                Console.WriteLine("  段落属性:");
                ExtractStyleParagraphProperties(style.StyleParagraphProperties, "    ");
            }

            // 文本属性
            if (style.StyleRunProperties != null)
            {
                Console.WriteLine("  文本属性:");
                ExtractStyleRunProperties(style.StyleRunProperties, "    ");
            }

            Console.WriteLine();
        }
    }

    static void ExtractMainDocumentFormat(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【3. 主文档内容格式】");
        Console.WriteLine(new string('-', 100));

        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null)
        {
            Console.WriteLine("文档主体为空");
            return;
        }

        int paraIndex = 0;
        int tableIndex = 0;

        foreach (var element in body.Elements())
        {
            if (element is Paragraph para)
            {
                paraIndex++;
                Console.WriteLine($"\n段落 #{paraIndex}:");
                Console.WriteLine(new string('~', 80));

                // 段落属性
                if (para.ParagraphProperties != null)
                {
                    Console.WriteLine("段落属性:");
                    ExtractParagraphProperties(para.ParagraphProperties, "  ");
                }

                // 段落文本内容
                var text = para.InnerText;
                Console.WriteLine($"文本内容: {(string.IsNullOrWhiteSpace(text) ? "[空段落]" : text)}");

                // 文本运行属性
                int runIndex = 0;
                foreach (var run in para.Elements<Run>())
                {
                    runIndex++;
                    Console.WriteLine($"\n  文本运行 #{runIndex}:");
                    Console.WriteLine($"    内容: {run.InnerText}");

                    if (run.RunProperties != null)
                    {
                        Console.WriteLine("    文本格式:");
                        ExtractRunProperties(run.RunProperties, "      ");
                    }
                }
            }
            else if (element is Table table)
            {
                tableIndex++;
                Console.WriteLine($"\n表格 #{tableIndex}:");
                Console.WriteLine(new string('~', 80));
                ExtractTableFormat(table);
            }
        }
    }

    static void ExtractParagraphProperties(ParagraphProperties props, string indent)
    {
        if (props.ParagraphStyleId != null)
            Console.WriteLine($"{indent}样式ID: {props.ParagraphStyleId.Val?.Value}");

        if (props.Justification != null)
            Console.WriteLine($"{indent}对齐方式: {props.Justification.Val?.Value}");

        if (props.Indentation != null)
        {
            var ind = props.Indentation;
            Console.WriteLine($"{indent}缩进:");
            if (ind.Left != null) Console.WriteLine($"{indent}  左缩进: {ind.Left.Value}");
            if (ind.Right != null) Console.WriteLine($"{indent}  右缩进: {ind.Right.Value}");
            if (ind.FirstLine != null) Console.WriteLine($"{indent}  首行缩进: {ind.FirstLine.Value}");
            if (ind.Hanging != null) Console.WriteLine($"{indent}  悬挂缩进: {ind.Hanging.Value}");
        }

        if (props.SpacingBetweenLines != null)
        {
            var spacing = props.SpacingBetweenLines;
            Console.WriteLine($"{indent}行间距:");
            if (spacing.Before != null) Console.WriteLine($"{indent}  段前间距: {spacing.Before.Value}");
            if (spacing.After != null) Console.WriteLine($"{indent}  段后间距: {spacing.After.Value}");
            if (spacing.Line != null) Console.WriteLine($"{indent}  行距: {spacing.Line.Value}");
            if (spacing.LineRule != null) Console.WriteLine($"{indent}  行距规则: {spacing.LineRule.Value}");
        }

        if (props.KeepNext != null)
            Console.WriteLine($"{indent}与下段同页: {props.KeepNext.Val?.Value ?? true}");

        if (props.KeepLines != null)
            Console.WriteLine($"{indent}段中不分页: {props.KeepLines.Val?.Value ?? true}");

        if (props.PageBreakBefore != null)
            Console.WriteLine($"{indent}段前分页: {props.PageBreakBefore.Val?.Value ?? true}");

        if (props.WidowControl != null)
            Console.WriteLine($"{indent}孤行控制: {props.WidowControl.Val?.Value ?? true}");

        if (props.SuppressLineNumbers != null)
            Console.WriteLine($"{indent}取消行号: {props.SuppressLineNumbers.Val?.Value ?? true}");

        if (props.Shading != null)
        {
            var shading = props.Shading;
            Console.WriteLine($"{indent}底纹:");
            if (shading.Fill != null) Console.WriteLine($"{indent}  填充色: #{shading.Fill.Value}");
            if (shading.Color != null) Console.WriteLine($"{indent}  前景色: #{shading.Color.Value}");
            if (shading.Val != null) Console.WriteLine($"{indent}  图案: {shading.Val.Value}");
        }

        if (props.ParagraphBorders != null)
        {
            Console.WriteLine($"{indent}边框:");
            ExtractBorders(props.ParagraphBorders, indent + "  ");
        }

        if (props.NumberingProperties != null)
        {
            var numProps = props.NumberingProperties;
            Console.WriteLine($"{indent}编号:");
            if (numProps.NumberingId != null) Console.WriteLine($"{indent}  编号ID: {numProps.NumberingId.Val?.Value}");
            if (numProps.NumberingLevelReference != null) Console.WriteLine($"{indent}  编号级别: {numProps.NumberingLevelReference.Val?.Value}");
        }
    }

    static void ExtractRunProperties(RunProperties props, string indent)
    {
        if (props.RunFonts != null)
        {
            var fonts = props.RunFonts;
            Console.WriteLine($"{indent}字体:");
            if (fonts.Ascii != null) Console.WriteLine($"{indent}  ASCII字体: {fonts.Ascii.Value}");
            if (fonts.EastAsia != null) Console.WriteLine($"{indent}  东亚字体: {fonts.EastAsia.Value}");
            if (fonts.HighAnsi != null) Console.WriteLine($"{indent}  高位ANSI字体: {fonts.HighAnsi.Value}");
            if (fonts.ComplexScript != null) Console.WriteLine($"{indent}  复杂字体: {fonts.ComplexScript.Value}");
        }

        if (props.FontSize != null)
            Console.WriteLine($"{indent}字号: {props.FontSize.Val?.Value} (半磅值)");

        if (props.FontSizeComplexScript != null)
            Console.WriteLine($"{indent}复杂字体字号: {props.FontSizeComplexScript.Val?.Value}");

        if (props.Bold != null)
            Console.WriteLine($"{indent}粗体: {props.Bold.Val?.Value ?? true}");

        if (props.BoldComplexScript != null)
            Console.WriteLine($"{indent}复杂字体粗体: {props.BoldComplexScript.Val?.Value ?? true}");

        if (props.Italic != null)
            Console.WriteLine($"{indent}斜体: {props.Italic.Val?.Value ?? true}");

        if (props.ItalicComplexScript != null)
            Console.WriteLine($"{indent}复杂字体斜体: {props.ItalicComplexScript.Val?.Value ?? true}");

        if (props.Underline != null)
        {
            Console.WriteLine($"{indent}下划线: {props.Underline.Val?.Value}");
            if (props.Underline.Color != null)
                Console.WriteLine($"{indent}  下划线颜色: #{props.Underline.Color.Value}");
        }

        if (props.Strike != null)
            Console.WriteLine($"{indent}删除线: {props.Strike.Val?.Value ?? true}");

        if (props.DoubleStrike != null)
            Console.WriteLine($"{indent}双删除线: {props.DoubleStrike.Val?.Value ?? true}");

        if (props.Color != null)
            Console.WriteLine($"{indent}文字颜色: #{props.Color.Val?.Value}");

        if (props.Highlight != null)
            Console.WriteLine($"{indent}高亮: {props.Highlight.Val?.Value}");

        if (props.Shading != null)
        {
            var shading = props.Shading;
            Console.WriteLine($"{indent}底纹:");
            if (shading.Fill != null) Console.WriteLine($"{indent}  填充色: #{shading.Fill.Value}");
            if (shading.Color != null) Console.WriteLine($"{indent}  前景色: #{shading.Color.Value}");
        }

        if (props.VerticalTextAlignment != null)
            Console.WriteLine($"{indent}垂直对齐: {props.VerticalTextAlignment.Val?.Value}");

        if (props.Spacing != null)
            Console.WriteLine($"{indent}字符间距: {props.Spacing.Val?.Value}");

        if (props.CharacterScale != null)
            Console.WriteLine($"{indent}字符缩放: {props.CharacterScale.Val?.Value}%");

        if (props.Position != null)
            Console.WriteLine($"{indent}位置: {props.Position.Val?.Value}");

        if (props.Kern != null)
            Console.WriteLine($"{indent}字距调整: {props.Kern.Val?.Value}");

        if (props.Emboss != null)
            Console.WriteLine($"{indent}阳文: {props.Emboss.Val?.Value ?? true}");

        if (props.Imprint != null)
            Console.WriteLine($"{indent}阴文: {props.Imprint.Val?.Value ?? true}");

        if (props.Shadow != null)
            Console.WriteLine($"{indent}阴影: {props.Shadow.Val?.Value ?? true}");

        if (props.Outline != null)
            Console.WriteLine($"{indent}轮廓: {props.Outline.Val?.Value ?? true}");

        if (props.SmallCaps != null)
            Console.WriteLine($"{indent}小型大写字母: {props.SmallCaps.Val?.Value ?? true}");

        if (props.Caps != null)
            Console.WriteLine($"{indent}全部大写: {props.Caps.Val?.Value ?? true}");

        if (props.Vanish != null)
            Console.WriteLine($"{indent}隐藏文字: {props.Vanish.Val?.Value ?? true}");
    }

    static void ExtractBorders(ParagraphBorders borders, string indent)
    {
        if (borders.TopBorder != null)
        {
            Console.WriteLine($"{indent}上边框:");
            ExtractBorderDetails(borders.TopBorder, indent + "  ");
        }
        if (borders.BottomBorder != null)
        {
            Console.WriteLine($"{indent}下边框:");
            ExtractBorderDetails(borders.BottomBorder, indent + "  ");
        }
        if (borders.LeftBorder != null)
        {
            Console.WriteLine($"{indent}左边框:");
            ExtractBorderDetails(borders.LeftBorder, indent + "  ");
        }
        if (borders.RightBorder != null)
        {
            Console.WriteLine($"{indent}右边框:");
            ExtractBorderDetails(borders.RightBorder, indent + "  ");
        }
    }

    static void ExtractBorderDetails(BorderType border, string indent)
    {
        if (border.Val != null) Console.WriteLine($"{indent}样式: {border.Val.Value}");
        if (border.Color != null) Console.WriteLine($"{indent}颜色: #{border.Color.Value}");
        if (border.Size != null) Console.WriteLine($"{indent}粗细: {border.Size.Value}");
        if (border.Space != null) Console.WriteLine($"{indent}间距: {border.Space.Value}");
    }

    static void ExtractTableFormat(Table table)
    {
        var tableProps = table.GetFirstChild<TableProperties>();
        if (tableProps != null)
        {
            Console.WriteLine("表格属性:");

            if (tableProps.TableStyle != null)
                Console.WriteLine($"  样式: {tableProps.TableStyle.Val?.Value}");

            if (tableProps.TableWidth != null)
                Console.WriteLine($"  宽度: {tableProps.TableWidth.Width?.Value} ({tableProps.TableWidth.Type?.Value})");

            if (tableProps.TableIndentation != null)
                Console.WriteLine($"  缩进: {tableProps.TableIndentation.Width?.Value}");

            if (tableProps.TableLayout != null)
                Console.WriteLine($"  布局: {tableProps.TableLayout.Type?.Value}");

            if (tableProps.TableBorders != null)
            {
                Console.WriteLine("  边框:");
                var borders = tableProps.TableBorders;
                if (borders.TopBorder != null) Console.WriteLine($"    上边框: {borders.TopBorder.Val?.Value}");
                if (borders.BottomBorder != null) Console.WriteLine($"    下边框: {borders.BottomBorder.Val?.Value}");
                if (borders.LeftBorder != null) Console.WriteLine($"    左边框: {borders.LeftBorder.Val?.Value}");
                if (borders.RightBorder != null) Console.WriteLine($"    右边框: {borders.RightBorder.Val?.Value}");
                if (borders.InsideHorizontalBorder != null) Console.WriteLine($"    内部横边框: {borders.InsideHorizontalBorder.Val?.Value}");
                if (borders.InsideVerticalBorder != null) Console.WriteLine($"    内部竖边框: {borders.InsideVerticalBorder.Val?.Value}");
            }

            if (tableProps.Shading != null)
            {
                var shading = tableProps.Shading;
                Console.WriteLine("  底纹:");
                if (shading.Fill != null) Console.WriteLine($"    填充色: #{shading.Fill.Value}");
            }
        }

        var rows = table.Elements<TableRow>().ToList();
        Console.WriteLine($"\n行数: {rows.Count}");

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            Console.WriteLine($"\n  行 #{i + 1}:");

            var rowProps = row.GetFirstChild<TableRowProperties>();
            if (rowProps != null)
            {
                var rowHeight = rowProps.GetFirstChild<TableRowHeight>();
                if (rowHeight != null)
                    Console.WriteLine($"    行高: {rowHeight.Val?.Value}");
                var tableHeader = rowProps.GetFirstChild<TableHeader>();
                if (tableHeader != null)
                    Console.WriteLine($"    标题行: true");
            }

            var cells = row.Elements<TableCell>().ToList();
            Console.WriteLine($"    单元格数: {cells.Count}");

            for (int j = 0; j < cells.Count; j++)
            {
                var cell = cells[j];
                Console.WriteLine($"\n    单元格 [{i + 1},{j + 1}]:");
                Console.WriteLine($"      内容: {cell.InnerText}");

                var cellProps = cell.GetFirstChild<TableCellProperties>();
                if (cellProps != null)
                {
                    if (cellProps.TableCellWidth != null)
                        Console.WriteLine($"      宽度: {cellProps.TableCellWidth.Width?.Value} ({cellProps.TableCellWidth.Type?.Value})");

                    if (cellProps.VerticalMerge != null)
                        Console.WriteLine($"      垂直合并: {cellProps.VerticalMerge.Val?.Value}");

                    if (cellProps.HorizontalMerge != null)
                        Console.WriteLine($"      水平合并: {cellProps.HorizontalMerge.Val?.Value}");

                    if (cellProps.Shading != null)
                    {
                        var shading = cellProps.Shading;
                        if (shading.Fill != null)
                            Console.WriteLine($"      背景色: #{shading.Fill.Value}");
                    }
                }
            }
        }
    }

    static void ExtractSections(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【4. 节信息】");
        Console.WriteLine(new string('-', 100));

        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return;

        var sections = body.Elements<SectionProperties>().ToList();

        // 获取最后一个段落的节属性
        var lastPara = body.Elements<Paragraph>().LastOrDefault();
        if (lastPara?.ParagraphProperties?.SectionProperties != null)
        {
            sections.Add(lastPara.ParagraphProperties.SectionProperties);
        }

        Console.WriteLine($"节数: {sections.Count}");

        for (int i = 0; i < sections.Count; i++)
        {
            var section = sections[i];
            Console.WriteLine($"\n节 #{i + 1}:");

            var pageSize = section.GetFirstChild<PageSize>();
            if (pageSize != null)
            {
                Console.WriteLine($"  页面大小:");
                Console.WriteLine($"    宽度: {pageSize.Width?.Value}");
                Console.WriteLine($"    高度: {pageSize.Height?.Value}");
                Console.WriteLine($"    方向: {pageSize.Orient?.Value}");
            }

            var margin = section.GetFirstChild<PageMargin>();
            if (margin != null)
            {
                Console.WriteLine($"  页边距:");
                if (margin.Top != null) Console.WriteLine($"    上: {margin.Top.Value}");
                if (margin.Bottom != null) Console.WriteLine($"    下: {margin.Bottom.Value}");
                if (margin.Left != null) Console.WriteLine($"    左: {margin.Left.Value}");
                if (margin.Right != null) Console.WriteLine($"    右: {margin.Right.Value}");
                if (margin.Header != null) Console.WriteLine($"    页眉距边界: {margin.Header.Value}");
                if (margin.Footer != null) Console.WriteLine($"    页脚距边界: {margin.Footer.Value}");
                if (margin.Gutter != null) Console.WriteLine($"    装订线: {margin.Gutter.Value}");
            }

            var columns = section.GetFirstChild<Columns>();
            if (columns != null)
            {
                Console.WriteLine($"  分栏:");
                if (columns.ColumnCount != null) Console.WriteLine($"    栏数: {columns.ColumnCount.Value}");
                if (columns.Space != null) Console.WriteLine($"    间距: {columns.Space.Value}");
                if (columns.Separator != null) Console.WriteLine($"    分隔线: {columns.Separator.Value}");
            }

            var sectionType = section.GetFirstChild<SectionType>();
            if (sectionType != null)
                Console.WriteLine($"  节类型: {sectionType.Val?.Value}");
        }
    }

    static void ExtractFonts(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【5. 字体表】");
        Console.WriteLine(new string('-', 100));

        var fontTablePart = doc.MainDocumentPart?.FontTablePart;
        if (fontTablePart?.Fonts == null)
        {
            Console.WriteLine("无字体表");
            return;
        }

        var fonts = fontTablePart.Fonts.Elements<DocumentFormat.OpenXml.Wordprocessing.Font>().ToList();
        Console.WriteLine($"字体数: {fonts.Count}");
        Console.WriteLine();

        foreach (var font in fonts)
        {
            Console.WriteLine($"字体名称: {font.Name?.Value}");

            // 简化字体属性提取，避免使用不存在的类型

            var fontFamily = font.GetFirstChild<FontFamily>();
            if (fontFamily != null)
                Console.WriteLine($"  字体族: {fontFamily.Val?.Value}");

            var pitch = font.GetFirstChild<Pitch>();
            if (pitch != null)
                Console.WriteLine($"  间距: {pitch.Val?.Value}");

            Console.WriteLine();
        }
    }

    static void ExtractNumbering(WordprocessingDocument doc)
    {
        Console.WriteLine("\n【6. 编号和列表】");
        Console.WriteLine(new string('-', 100));

        var numberingPart = doc.MainDocumentPart?.NumberingDefinitionsPart;
        if (numberingPart?.Numbering == null)
        {
            Console.WriteLine("无编号定义");
            return;
        }

        var abstractNums = numberingPart.Numbering.Elements<AbstractNum>().ToList();
        Console.WriteLine($"抽象编号定义数: {abstractNums.Count}");

        foreach (var abstractNum in abstractNums)
        {
            Console.WriteLine($"\n抽象编号ID: {abstractNum.AbstractNumberId?.Value}");

            var levels = abstractNum.Elements<Level>().ToList();
            Console.WriteLine($"  级别数: {levels.Count}");

            foreach (var level in levels)
            {
                Console.WriteLine($"\n  级别 {level.LevelIndex?.Value}:");

                if (level.NumberingFormat != null)
                    Console.WriteLine($"    格式: {level.NumberingFormat.Val?.Value}");

                if (level.LevelText != null)
                    Console.WriteLine($"    文本: {level.LevelText.Val?.Value}");

                if (level.StartNumberingValue != null)
                    Console.WriteLine($"    起始值: {level.StartNumberingValue.Val?.Value}");

                if (level.LevelJustification != null)
                    Console.WriteLine($"    对齐: {level.LevelJustification.Val?.Value}");

                if (level.PreviousParagraphProperties != null)
                {
                    Console.WriteLine($"    段落属性:");
                    ExtractLevelParagraphProperties(level.PreviousParagraphProperties, "      ");
                }

                if (level.NumberingSymbolRunProperties != null)
                {
                    Console.WriteLine($"    编号文本属性:");
                    ExtractNumberingRunProperties(level.NumberingSymbolRunProperties, "      ");
                }
            }
        }

        var numberingInstances = numberingPart.Numbering.Elements<NumberingInstance>().ToList();
        Console.WriteLine($"\n\n编号实例数: {numberingInstances.Count}");

        foreach (var numInstance in numberingInstances)
        {
            Console.WriteLine($"\n编号ID: {numInstance.NumberID?.Value}");

            if (numInstance.AbstractNumId != null)
                Console.WriteLine($"  基于抽象编号: {numInstance.AbstractNumId.Val?.Value}");
        }
    }

    static void ExtractStyleParagraphProperties(StyleParagraphProperties props, string indent)
    {
        if (props.Justification != null)
            Console.WriteLine($"{indent}对齐方式: {props.Justification.Val?.Value}");

        if (props.Indentation != null)
        {
            var ind = props.Indentation;
            Console.WriteLine($"{indent}缩进:");
            if (ind.Left != null) Console.WriteLine($"{indent}  左缩进: {ind.Left.Value}");
            if (ind.Right != null) Console.WriteLine($"{indent}  右缩进: {ind.Right.Value}");
            if (ind.FirstLine != null) Console.WriteLine($"{indent}  首行缩进: {ind.FirstLine.Value}");
            if (ind.Hanging != null) Console.WriteLine($"{indent}  悬挂缩进: {ind.Hanging.Value}");
        }

        if (props.SpacingBetweenLines != null)
        {
            var spacing = props.SpacingBetweenLines;
            Console.WriteLine($"{indent}行间距:");
            if (spacing.Before != null) Console.WriteLine($"{indent}  段前间距: {spacing.Before.Value}");
            if (spacing.After != null) Console.WriteLine($"{indent}  段后间距: {spacing.After.Value}");
            if (spacing.Line != null) Console.WriteLine($"{indent}  行距: {spacing.Line.Value}");
            if (spacing.LineRule != null) Console.WriteLine($"{indent}  行距规则: {spacing.LineRule.Value}");
        }
    }

    static void ExtractStyleRunProperties(StyleRunProperties props, string indent)
    {
        if (props.RunFonts != null)
        {
            var fonts = props.RunFonts;
            Console.WriteLine($"{indent}字体:");
            if (fonts.Ascii != null) Console.WriteLine($"{indent}  ASCII字体: {fonts.Ascii.Value}");
            if (fonts.EastAsia != null) Console.WriteLine($"{indent}  东亚字体: {fonts.EastAsia.Value}");
            if (fonts.HighAnsi != null) Console.WriteLine($"{indent}  高位ANSI字体: {fonts.HighAnsi.Value}");
            if (fonts.ComplexScript != null) Console.WriteLine($"{indent}  复杂字体: {fonts.ComplexScript.Value}");
        }

        if (props.FontSize != null)
            Console.WriteLine($"{indent}字号: {props.FontSize.Val?.Value} (半磅值)");

        if (props.Bold != null)
            Console.WriteLine($"{indent}粗体: {props.Bold.Val?.Value ?? true}");

        if (props.Italic != null)
            Console.WriteLine($"{indent}斜体: {props.Italic.Val?.Value ?? true}");

        if (props.Color != null)
            Console.WriteLine($"{indent}文字颜色: #{props.Color.Val?.Value}");
    }

    static void ExtractLevelParagraphProperties(PreviousParagraphProperties props, string indent)
    {
        if (props.Indentation != null)
        {
            var ind = props.Indentation;
            Console.WriteLine($"{indent}缩进:");
            if (ind.Left != null) Console.WriteLine($"{indent}  左缩进: {ind.Left.Value}");
            if (ind.Right != null) Console.WriteLine($"{indent}  右缩进: {ind.Right.Value}");
            if (ind.FirstLine != null) Console.WriteLine($"{indent}  首行缩进: {ind.FirstLine.Value}");
            if (ind.Hanging != null) Console.WriteLine($"{indent}  悬挂缩进: {ind.Hanging.Value}");
        }

        if (props.SpacingBetweenLines != null)
        {
            var spacing = props.SpacingBetweenLines;
            Console.WriteLine($"{indent}行间距:");
            if (spacing.Before != null) Console.WriteLine($"{indent}  段前间距: {spacing.Before.Value}");
            if (spacing.After != null) Console.WriteLine($"{indent}  段后间距: {spacing.After.Value}");
        }
    }

    static void ExtractNumberingRunProperties(NumberingSymbolRunProperties props, string indent)
    {
        if (props.RunFonts != null)
        {
            var fonts = props.RunFonts;
            Console.WriteLine($"{indent}字体:");
            if (fonts.Ascii != null) Console.WriteLine($"{indent}  ASCII字体: {fonts.Ascii.Value}");
            if (fonts.EastAsia != null) Console.WriteLine($"{indent}  东亚字体: {fonts.EastAsia.Value}");
        }

        if (props.FontSize != null)
            Console.WriteLine($"{indent}字号: {props.FontSize.Val?.Value} (半磅值)");

        if (props.Bold != null)
            Console.WriteLine($"{indent}粗体: {props.Bold.Val?.Value ?? true}");

        if (props.Color != null)
            Console.WriteLine($"{indent}文字颜色: #{props.Color.Val?.Value}");
    }
}
