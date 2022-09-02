package com.money.docx.item;

import com.money.docx.style.DocxStyle;
import lombok.Getter;
import lombok.Setter;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯项
 * @createTime : 2022-08-20 15:47:02
 */
@Getter
@Setter
public abstract class DocxItem {

    /**
     * 多克斯风格
     */
    private DocxStyle style;

    public DocxItem() {
    }

    public DocxItem(DocxStyle style) {
        this.style = style;
    }
}
