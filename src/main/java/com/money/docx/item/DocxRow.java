package com.money.docx.item;

import com.money.docx.style.DocxStyle;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;

import java.util.List;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯行
 * @createTime : 2022-08-20 15:50:53
 */
@Getter
@Setter
@RequiredArgsConstructor
public class DocxRow extends DocxItem {

    /**
     * 单元格
     */
    private final List<DocxCell> cells;

    public DocxRow(List<DocxCell> cells, DocxStyle style) {
        super(style);
        this.cells = cells;
    }
}
