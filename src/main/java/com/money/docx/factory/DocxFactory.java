package com.money.docx.factory;

import com.money.docx.item.*;
import com.money.docx.style.DocxStyle;
import lombok.NonNull;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.*;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯工厂
 * <a href="http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/index.html">docx4j支持的word标签</a>
 * 可以在该文档查看相关样式的xml标签，docx4j都有对应的类，按其嵌套方式去组装就能有相同效果
 * @createTime : 2022-08-20 15:58:01
 */
public class DocxFactory extends StyleFactory {

    /**
     * 新页面
     *
     * @return {@link P}
     */
    public static P newPage() {
        P p = factory.createP();
        R r = factory.createR();
        Br br = factory.createBr();
        br.setType(STBrType.PAGE);
        r.getContent().add(br);
        p.getContent().add(r);
        return p;
    }

    /**
     * 创建标题
     *
     * @param styleId  风格id(需要word的标题样式)
     * @param docxText 多克斯文本
     * @return {@link P}
     */
    public static P createHeading(String styleId, DocxText docxText) {
        P p = factory.createP();
        PPr pPr = factory.createPPr();
        p.setPPr(pPr);
        // 标题级别
        PPrBase.PStyle pStyle = factory.createPPrBasePStyle();
        pStyle.setVal(styleId);
        pPr.setPStyle(pStyle);
        //项目符号
        PPrBase.NumPr.Ilvl ilvl = factory.createPPrBaseNumPrIlvl();
        ilvl.setVal(BigInteger.ZERO);
        //编号
        PPrBase.NumPr.NumId numId = factory.createPPrBaseNumPrNumId();
        numId.setVal(BigInteger.ZERO);
        PPrBase.NumPr numPr = factory.createPPrBaseNumPr();
        numPr.setIlvl(ilvl);
        numPr.setNumId(numId);
        pPr.setNumPr(numPr);
        R r = createText(docxText.getText(), docxText.getStyle());
        p.getContent().add(r);
        return p;
    }

    /**
     * 创建段落
     *
     * @param docxText 多克斯文本
     * @return {@link P}
     */
    public static P createParagraph(DocxText docxText) {
        P p = addParagraphStyle(factory.createP(), docxText.getStyle());
        R r = createText(docxText.getText(), docxText.getStyle());
        p.getContent().add(r);
        return p;
    }

    /**
     * 创建段落
     *
     * @param docxImage 多克斯形象
     * @param wpk       劳动党
     * @return {@link P}
     */
    public static P createParagraph(DocxImage docxImage, WordprocessingMLPackage wpk) {
        P p = addParagraphStyle(factory.createP(), docxImage.getStyle());
        R r = createImage(docxImage, wpk);
        p.getContent().add(r);
        return p;
    }

    /**
     * 创建表格
     *
     * @param docxTable 多克斯表
     * @param wpk       劳动党
     * @return {@link Tbl}
     */
    public static Tbl createTable(DocxTable docxTable, WordprocessingMLPackage wpk) {
        if (docxTable.getColumnCount() == 0) {
            return null;
        }
        // 创建空表
        int writableWidthTwips = wpk.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
        Tbl tbl = TblFactory.createTable(docxTable.getRowCount(), docxTable.getColumnCount(), writableWidthTwips / docxTable.getColumnCount());
        // 调整列宽比例
        adjustTableColWidth(tbl, docxTable.getRows().get(0));
        // 填充表
        for (int i = 0; i < docxTable.getRowCount(); i++) {
            DocxRow docxRow = docxTable.getRows().get(i);
            Tr tr = (Tr) tbl.getContent().get(i);
            for (int j = 0; j < docxTable.getColumnCount(); j++) {
                DocxStyle style = DocxStyle.builder().build();
                style.merge(docxRow.getStyle());
                DocxCell cell = docxRow.getCells().get(j);
                style.merge(cell.getStyle());
                Tc tc = addCellStyle((Tc) tr.getContent().get(j), style);
                tc.getContent().clear();
                //复杂单元格
                if (cell.getPList() != null) {
                    tc.getContent().addAll(cell.getPList().stream().map(p -> addParagraphStyle(p, style)).collect(Collectors.toList()));
                } else {
                    //普通单元格
                    P p = factory.createP();
                    if (cell.getDocxText() != null) {
                        DocxText docxText = cell.getDocxText();
                        style.merge(docxText.getStyle());
                        docxText.setStyle(style);
                        p = createParagraph(docxText);
                    } else if (cell.getDocxImage() != null) {
                        DocxImage docxImage = cell.getDocxImage();
                        style.merge(docxImage.getStyle());
                        docxImage.setStyle(style);
                        p = createParagraph(docxImage, wpk);
                    }
                    tc.getContent().add(p);
                }
            }
        }
        return tbl;
    }

    /**
     * 调整表格列宽
     *
     * @param firstRow           第一行
     */
    private static void adjustTableColWidth(Tbl tbl, DocxRow firstRow) {
        CTTblLayoutType ctTblLayoutType = factory.createCTTblLayoutType();
        ctTblLayoutType.setType(STTblLayoutType.FIXED);
        tbl.getTblPr().setTblLayout(ctTblLayoutType);

        TblGrid tblGrid = tbl.getTblGrid();
        BigDecimal gridColSum = BigDecimal.valueOf(tblGrid.getGridCol().stream().mapToLong(gc -> gc.getW().longValue()).sum());

        List<DocxCell> cells = firstRow.getCells();
        List<BigDecimal> cellWidthList = new ArrayList<>(cells.size());
        cells.forEach(cell -> {
            if (cell.getStyle() != null && cell.getStyle().getCellWidth() != null) {
                cellWidthList.add(BigDecimal.valueOf(cell.getStyle().getCellWidth()));
            } else {
                cellWidthList.add(BigDecimal.valueOf(1000L));
            }
        });
        BigDecimal sum = BigDecimal.valueOf(cellWidthList.stream().mapToLong(BigDecimal::longValue).sum());

        for (int i = 0; i < cells.size(); i++) {
            BigDecimal colWidth = cellWidthList.get(i).divide(sum, 3, RoundingMode.HALF_UP).multiply(gridColSum);
            tblGrid.getGridCol().get(i).setW(colWidth.toBigInteger());
        }
    }

    /**
     * 创建文本
     *
     * @param textToInsert 文本插入
     * @param style        多克斯风格
     * @return {@link R}
     */
    public static R createText(String textToInsert, DocxStyle style) {
        R r = addTextStyle(factory.createR(), style);
        String[] textArr = textToInsert.split("<br>");
        for (String t : textArr) {
            Text text = factory.createText();
            text.setValue(t);
            r.getContent().add(text);
            Br br = factory.createBr();
            br.setType(STBrType.TEXT_WRAPPING);
            r.getContent().add(br);
        }
        r.getContent().remove(r.getContent().size() - 1);
        return r;
    }

    /**
     * 创建图像
     *
     * @param docxImage 多克斯形象
     * @param wpk       劳动党
     * @return {@link R}
     */
    public static R createImage(DocxImage docxImage, WordprocessingMLPackage wpk) {
        R r = factory.createR();
        try {
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wpk, docxImage.getBytes());
            Inline inline;
            if (docxImage.getWidth() != null) {
                inline = imagePart.createImageInline(docxImage.getFileNameHint(), docxImage.getAltText(), 1, 1, docxImage.getWidth(), false);
            } else {
                inline = imagePart.createImageInline(docxImage.getFileNameHint(), docxImage.getAltText(), 1, 1, false);
            }
            Drawing drawing = factory.createDrawing();
            r.getContent().add(drawing);
            drawing.getAnchorOrInline().add(inline);
        } catch (Exception e) {
            throw new RuntimeException("wpk create image fail");
        }
        return r;
    }

    // =========================== tool

    public static byte[] readImage(@NonNull InputStream in, boolean close) {
        byte[] result;
        try {
            int available = in.available();
            result = new byte[available];
            int readLength = in.read(result);
            if (readLength != available) {
                throw new IOException("File length is [" + available + " ] but read [" + readLength + "]!");
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (close) {
                try {
                    in.close();
                } catch (Exception ignored) {
                }
            }
        }
        return result;
    }
}
