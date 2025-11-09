package com.docx4j.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.*;
import org.docx4j.XmlUtils;
import org.docx4j.model.structure.SectionWrapper;

import jakarta.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

/**
 * Comprehensive DOCX format extractor using docx4j
 */
public class DocxFormatExtractor {

    public static void main(String[] args) {
        try {
            String docxPath = "test.docx";
            File docxFile = new File(docxPath);

            if (!docxFile.exists()) {
                System.err.println("文件不存在: " + docxPath);
                return;
            }

            System.out.println("===============================================");
            System.out.println("    DOCX 格式全面提取分析");
            System.out.println("    文件: " + docxPath);
            System.out.println("===============================================\n");

            // Load the document
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxFile);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            // 1. Extract document properties
            extractDocumentProperties(wordMLPackage);

            // 2. Extract page setup
            extractPageSetup(mainDocumentPart);

            // 3. Extract styles
            extractStyles(wordMLPackage);

            // 4. Extract paragraph and text formatting
            extractContentFormatting(mainDocumentPart);

            // 5. Extract tables
            extractTables(mainDocumentPart);

            // 6. Extract numbering and lists
            extractNumbering(wordMLPackage);

            System.out.println("\n===============================================");
            System.out.println("    提取完成!");
            System.out.println("===============================================");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Extract document properties
     */
    private static void extractDocumentProperties(WordprocessingMLPackage wordMLPackage) {
        System.out.println("\n【1. 文档属性】");
        System.out.println("----------------------------------------");
        try {
            if (wordMLPackage.getDocPropsCorePart() != null) {
                org.docx4j.docProps.core.CoreProperties coreProps =
                    (org.docx4j.docProps.core.CoreProperties) wordMLPackage.getDocPropsCorePart().getContents();

                if (coreProps.getTitle() != null) {
                    System.out.println("标题: " + coreProps.getTitle());
                }
                if (coreProps.getCreator() != null) {
                    System.out.println("作者: " + coreProps.getCreator());
                }
                if (coreProps.getSubject() != null) {
                    System.out.println("主题: " + coreProps.getSubject());
                }
                if (coreProps.getDescription() != null) {
                    System.out.println("描述: " + coreProps.getDescription());
                }
                if (coreProps.getCreated() != null) {
                    System.out.println("创建时间: " + coreProps.getCreated());
                }
                if (coreProps.getModified() != null) {
                    System.out.println("修改时间: " + coreProps.getModified());
                }
            }
        } catch (Exception e) {
            System.out.println("无法提取文档属性: " + e.getMessage());
        }
    }

    /**
     * Extract page setup information
     */
    private static void extractPageSetup(MainDocumentPart mainDocumentPart) {
        System.out.println("\n【2. 页面设置】");
        System.out.println("----------------------------------------");
        try {
            // Get the section properties from the body
            SectPr sectPr = mainDocumentPart.getJaxbElement().getBody().getSectPr();

            if (sectPr != null) {
                int i = 0;

                System.out.println("节 " + (i + 1) + ":");

                if (sectPr != null) {
                    // Page size
                    SectPr.PgSz pgSz = sectPr.getPgSz();
                    if (pgSz != null) {
                        System.out.println("  页面尺寸: " + twipsToMM(pgSz.getW()) + "mm × " +
                                         twipsToMM(pgSz.getH()) + "mm");
                        if (pgSz.getOrient() != null) {
                            System.out.println("  页面方向: " + pgSz.getOrient());
                        }
                    }

                    // Page margins
                    SectPr.PgMar pgMar = sectPr.getPgMar();
                    if (pgMar != null) {
                        System.out.println("  页边距:");
                        System.out.println("    上: " + twipsToMM(pgMar.getTop()) + "mm");
                        System.out.println("    下: " + twipsToMM(pgMar.getBottom()) + "mm");
                        System.out.println("    左: " + twipsToMM(pgMar.getLeft()) + "mm");
                        System.out.println("    右: " + twipsToMM(pgMar.getRight()) + "mm");
                    }

                    // Columns
                    CTColumns cols = sectPr.getCols();
                    if (cols != null && cols.getNum() != null) {
                        System.out.println("  栏数: " + cols.getNum());
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("无法提取页面设置: " + e.getMessage());
        }
    }

    /**
     * Extract styles
     */
    private static void extractStyles(WordprocessingMLPackage wordMLPackage) {
        System.out.println("\n【3. 样式定义】");
        System.out.println("----------------------------------------");
        try {
            StyleDefinitionsPart stylesPart = wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylesPart != null) {
                Styles styles = stylesPart.getContents();
                List<Style> styleList = styles.getStyle();

                System.out.println("共有 " + styleList.size() + " 个样式定义\n");

                int count = 0;
                for (Style style : styleList) {
                    if (count++ > 10) { // Limit output to first 10 styles
                        System.out.println("... (" + (styleList.size() - 10) + " 个更多样式)");
                        break;
                    }

                    System.out.println("样式 ID: " + style.getStyleId());
                    if (style.getName() != null) {
                        System.out.println("  名称: " + style.getName().getVal());
                    }
                    if (style.getType() != null) {
                        System.out.println("  类型: " + style.getType());
                    }

                    // Paragraph properties
                    if (style.getPPr() != null) {
                        PPr pPr = style.getPPr();
                        extractParagraphProperties("  ", pPr);
                    }

                    // Character properties
                    if (style.getRPr() != null) {
                        RPr rPr = style.getRPr();
                        extractRunProperties("  ", rPr);
                    }
                    System.out.println();
                }
            }
        } catch (Exception e) {
            System.out.println("无法提取样式: " + e.getMessage());
        }
    }

    /**
     * Extract content formatting (paragraphs and runs)
     */
    private static void extractContentFormatting(MainDocumentPart mainDocumentPart) {
        System.out.println("\n【4. 内容格式】");
        System.out.println("----------------------------------------");
        try {
            List<Object> content = mainDocumentPart.getContent();
            int paraCount = 0;

            for (Object obj : content) {
                if (obj instanceof P) {
                    P paragraph = (P) obj;
                    paraCount++;

                    if (paraCount > 20) { // Limit to first 20 paragraphs
                        System.out.println("\n... (更多段落省略)");
                        break;
                    }

                    System.out.println("\n段落 " + paraCount + ":");

                    // Get paragraph text
                    String text = extractText(paragraph);
                    if (text != null && !text.trim().isEmpty()) {
                        System.out.println("  内容: " + (text.length() > 50 ? text.substring(0, 50) + "..." : text));
                    }

                    // Paragraph properties
                    PPr pPr = paragraph.getPPr();
                    if (pPr != null) {
                        extractParagraphProperties("  ", pPr);
                    }

                    // Run properties
                    List<Object> runs = paragraph.getContent();
                    for (Object runObj : runs) {
                        if (runObj instanceof R) {
                            R run = (R) runObj;
                            RPr rPr = run.getRPr();
                            if (rPr != null) {
                                extractRunProperties("    ", rPr);
                                break; // Only show first run's properties
                            }
                        }
                    }
                }
            }

            System.out.println("\n总段落数: " + countParagraphs(content));
        } catch (Exception e) {
            System.out.println("无法提取内容格式: " + e.getMessage());
        }
    }

    /**
     * Extract paragraph properties
     */
    private static void extractParagraphProperties(String indent, PPr pPr) {
        if (pPr == null) return;

        // Style
        if (pPr.getPStyle() != null && pPr.getPStyle().getVal() != null) {
            System.out.println(indent + "样式: " + pPr.getPStyle().getVal());
        }

        // Alignment
        if (pPr.getJc() != null && pPr.getJc().getVal() != null) {
            System.out.println(indent + "对齐: " + pPr.getJc().getVal());
        }

        // Indentation
        if (pPr.getInd() != null) {
            PPrBase.Ind ind = pPr.getInd();
            if (ind.getLeft() != null) {
                System.out.println(indent + "左缩进: " + twipsToMM(ind.getLeft()) + "mm");
            }
            if (ind.getRight() != null) {
                System.out.println(indent + "右缩进: " + twipsToMM(ind.getRight()) + "mm");
            }
            if (ind.getFirstLine() != null) {
                System.out.println(indent + "首行缩进: " + twipsToMM(ind.getFirstLine()) + "mm");
            }
            if (ind.getHanging() != null) {
                System.out.println(indent + "悬挂缩进: " + twipsToMM(ind.getHanging()) + "mm");
            }
        }

        // Spacing
        if (pPr.getSpacing() != null) {
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing.getBefore() != null) {
                System.out.println(indent + "段前间距: " + twipsToPoint(spacing.getBefore()) + "磅");
            }
            if (spacing.getAfter() != null) {
                System.out.println(indent + "段后间距: " + twipsToPoint(spacing.getAfter()) + "磅");
            }
            if (spacing.getLine() != null) {
                System.out.println(indent + "行距: " + spacing.getLine() +
                                 (spacing.getLineRule() != null ? " (" + spacing.getLineRule() + ")" : ""));
            }
        }
    }

    /**
     * Extract run (character) properties
     */
    private static void extractRunProperties(String indent, RPr rPr) {
        if (rPr == null) return;

        System.out.println(indent + "字符格式:");

        // Font
        if (rPr.getRFonts() != null) {
            RFonts fonts = rPr.getRFonts();
            if (fonts.getAscii() != null) {
                System.out.println(indent + "  字体(ASCII): " + fonts.getAscii());
            }
            if (fonts.getEastAsia() != null) {
                System.out.println(indent + "  字体(东亚): " + fonts.getEastAsia());
            }
        }

        // Font size
        if (rPr.getSz() != null && rPr.getSz().getVal() != null) {
            System.out.println(indent + "  字号: " + (rPr.getSz().getVal().intValue() / 2) + "磅");
        }

        // Color
        if (rPr.getColor() != null && rPr.getColor().getVal() != null) {
            System.out.println(indent + "  颜色: #" + rPr.getColor().getVal());
        }

        // Bold
        if (rPr.getB() != null) {
            System.out.println(indent + "  加粗: " + rPr.getB().isVal());
        }

        // Italic
        if (rPr.getI() != null) {
            System.out.println(indent + "  斜体: " + rPr.getI().isVal());
        }

        // Underline
        if (rPr.getU() != null && rPr.getU().getVal() != null) {
            System.out.println(indent + "  下划线: " + rPr.getU().getVal());
        }

        // Highlight
        if (rPr.getHighlight() != null && rPr.getHighlight().getVal() != null) {
            System.out.println(indent + "  高亮: " + rPr.getHighlight().getVal());
        }
    }

    /**
     * Extract tables
     */
    private static void extractTables(MainDocumentPart mainDocumentPart) {
        System.out.println("\n【5. 表格】");
        System.out.println("----------------------------------------");
        try {
            List<Object> content = mainDocumentPart.getContent();
            int tableCount = 0;

            for (Object obj : content) {
                if (obj instanceof Tbl) {
                    Tbl table = (Tbl) obj;
                    tableCount++;

                    System.out.println("\n表格 " + tableCount + ":");

                    // Table properties
                    if (table.getTblPr() != null) {
                        TblPr tblPr = table.getTblPr();

                        if (tblPr.getTblW() != null) {
                            System.out.println("  宽度: " + tblPr.getTblW().getW() + " (" + tblPr.getTblW().getType() + ")");
                        }

                        if (tblPr.getJc() != null && tblPr.getJc().getVal() != null) {
                            System.out.println("  对齐: " + tblPr.getJc().getVal());
                        }
                    }

                    // Count rows and columns
                    List<Object> rows = table.getContent();
                    int rowCount = 0;
                    int maxCols = 0;

                    for (Object rowObj : rows) {
                        if (rowObj instanceof Tr) {
                            rowCount++;
                            Tr row = (Tr) rowObj;
                            int colCount = 0;
                            for (Object cellObj : row.getContent()) {
                                if (cellObj instanceof Tc) {
                                    colCount++;
                                }
                            }
                            maxCols = Math.max(maxCols, colCount);
                        }
                    }

                    System.out.println("  行数: " + rowCount);
                    System.out.println("  列数: " + maxCols);
                }
            }

            System.out.println("\n总表格数: " + tableCount);
        } catch (Exception e) {
            System.out.println("无法提取表格: " + e.getMessage());
        }
    }

    /**
     * Extract numbering and lists
     */
    private static void extractNumbering(WordprocessingMLPackage wordMLPackage) {
        System.out.println("\n【6. 编号和列表】");
        System.out.println("----------------------------------------");
        try {
            if (wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart() != null) {
                Numbering numbering = wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart().getContents();

                if (numbering.getNum() != null) {
                    System.out.println("编号实例数: " + numbering.getNum().size());
                }

                if (numbering.getAbstractNum() != null) {
                    System.out.println("抽象编号定义数: " + numbering.getAbstractNum().size());

                    for (int i = 0; i < Math.min(3, numbering.getAbstractNum().size()); i++) {
                        Numbering.AbstractNum abstractNum = numbering.getAbstractNum().get(i);
                        System.out.println("\n抽象编号 " + abstractNum.getAbstractNumId() + ":");

                        if (abstractNum.getLvl() != null) {
                            System.out.println("  级别数: " + abstractNum.getLvl().size());
                            for (Lvl lvl : abstractNum.getLvl()) {
                                System.out.println("    级别 " + lvl.getIlvl() + ": " +
                                                 (lvl.getNumFmt() != null ? lvl.getNumFmt().getVal() : ""));
                            }
                        }
                    }
                }
            } else {
                System.out.println("文档中没有编号定义");
            }
        } catch (Exception e) {
            System.out.println("无法提取编号: " + e.getMessage());
        }
    }

    /**
     * Extract text from paragraph
     */
    private static String extractText(P paragraph) {
        StringBuilder text = new StringBuilder();
        for (Object obj : paragraph.getContent()) {
            if (obj instanceof R) {
                R run = (R) obj;
                for (Object content : run.getContent()) {
                    if (content instanceof JAXBElement) {
                        JAXBElement<?> element = (JAXBElement<?>) content;
                        if (element.getValue() instanceof Text) {
                            Text t = (Text) element.getValue();
                            text.append(t.getValue());
                        }
                    } else if (content instanceof Text) {
                        Text t = (Text) content;
                        text.append(t.getValue());
                    }
                }
            }
        }
        return text.toString();
    }

    /**
     * Count total paragraphs
     */
    private static int countParagraphs(List<Object> content) {
        int count = 0;
        for (Object obj : content) {
            if (obj instanceof P) {
                count++;
            }
        }
        return count;
    }

    /**
     * Convert twips to millimeters
     */
    private static double twipsToMM(Object twips) {
        if (twips == null) return 0;
        double value = 0;
        if (twips instanceof Integer) {
            value = ((Integer) twips).doubleValue();
        } else if (twips instanceof Long) {
            value = ((Long) twips).doubleValue();
        } else if (twips instanceof java.math.BigInteger) {
            value = ((java.math.BigInteger) twips).doubleValue();
        }
        return Math.round(value / 56.7 * 10) / 10.0; // 1 twip = 1/1440 inch, 1 inch = 25.4mm
    }

    /**
     * Convert twips to points
     */
    private static double twipsToPoint(Object twips) {
        if (twips == null) return 0;
        double value = 0;
        if (twips instanceof Integer) {
            value = ((Integer) twips).doubleValue();
        } else if (twips instanceof Long) {
            value = ((Long) twips).doubleValue();
        } else if (twips instanceof java.math.BigInteger) {
            value = ((java.math.BigInteger) twips).doubleValue();
        }
        return Math.round(value / 20.0 * 10) / 10.0; // 1 point = 20 twips
    }
}
