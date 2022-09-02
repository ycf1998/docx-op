package com.money.docx.item;

import com.money.docx.style.DocxStyle;
import lombok.Getter;
import org.docx4j.wml.P;

import java.util.List;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯单元格
 * @createTime : 2022-08-27 11:58:07
 */
@Getter
public class DocxCell extends DocxItem {

    /**
     * p列表
     */
    private List<P> pList;

    /**
     * 多克斯文本
     */
    private DocxText docxText;

    /**
     * 多克斯形象
     */
    private DocxImage docxImage;

    public DocxCell(List<P> pList) {
        this.pList = pList;
    }

    public DocxCell(DocxText docxText) {
        this.docxText = docxText;
    }

    public DocxCell(DocxImage docxImage) {
        this.docxImage = docxImage;
    }

    public DocxCell(List<P> pList, DocxStyle style) {
        super(style);
        this.pList = pList;
    }

    public DocxCell(DocxText docxText, DocxStyle style) {
        super(style);
        this.docxText = docxText;
    }

    public DocxCell(DocxImage docxImage, DocxStyle style) {
        super(style);
        this.docxImage = docxImage;
    }
}
