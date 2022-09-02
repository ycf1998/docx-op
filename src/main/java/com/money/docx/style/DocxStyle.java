package com.money.docx.style;

import com.money.docx.constant.TbMergeEnum;
import lombok.Builder;
import lombok.Data;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.STVerticalJc;

import java.util.List;


@Data
@Builder
public class DocxStyle {

    // =============== 段落

    /**
     * 对齐
     */
    private JcEnumeration align;

    // =============== 表格

    /**
     * 外边框 上右下左
     */
    private List<Long> margins;

    /**
     * 单元格颜色
     */
    private String cellColor;

    /**
     * 垂直对齐
     */
    private STVerticalJc vAlign;

    /**
     * 垂直合并
     */
    private TbMergeEnum vMerge;

    /**
     * 水平合并
     */
    private TbMergeEnum hMerge;

    // =============== 文本

    /**
     * 粗体
     */
    private Boolean bold;

    /**
     * 斜体
     */
    private Boolean italic;

    /**
     * 下划线
     */
    private Boolean underline;

    /**
     * 字体大小
     */
    private Long fontSize;

    /**
     * 字体颜色
     */
    private String fontColor;

    /**
     * 字体
     */
    private String fontFamily;

    /**
     * 合并
     *
     * @param style 风格
     */
    public void merge(DocxStyle style) {
        if (style == null) {
            return;
        }
        if (style.align != null) {
            this.align = style.align;
        }
        if (style.margins != null) {
            this.margins = style.margins;
        }
        if (style.cellColor != null) {
            this.cellColor = style.cellColor;
        }
        if (style.vAlign != null) {
            this.vAlign = style.vAlign;
        }
        if (style.vMerge != null) {
            this.vMerge = style.vMerge;
        }
        if (style.hMerge != null) {
            this.hMerge = style.hMerge;
        }
        if (style.bold != null) {
            this.bold = style.bold;
        }
        if (style.italic != null) {
            this.italic = style.italic;
        }
        if (style.underline != null) {
            this.underline = style.underline;
        }
        if (style.fontSize != null) {
            this.fontSize = style.fontSize;
        }
        if (style.fontColor != null) {
            this.fontColor = style.fontColor;
        }
        if (style.fontFamily != null) {
            this.fontFamily = style.fontFamily;
        }
    }
}
