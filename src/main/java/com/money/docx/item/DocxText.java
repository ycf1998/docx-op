package com.money.docx.item;

import com.money.docx.style.DocxStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
public class DocxText extends DocxItem {

    /**
     * 文本，支持<br>换行
     */
    private String text;

    public DocxText(String text, DocxStyle style) {
        super(style);
        this.text = text;
    }
}
