package com.money.docx.item;

import com.money.docx.style.DocxStyle;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯图片
 * @createTime : 2022-08-20 15:54:22
 */
@Getter
@Setter
@RequiredArgsConstructor
public class DocxImage extends DocxItem {

    /**
     * 字节
     */
    private final byte[] bytes;

    /**
     * 文件名提示
     */
    private String fileNameHint = "";

    /**
     * alt文本
     */
    private String altText = "";

    /**
     * 宽度
     */
    private Long width;


    public DocxImage(byte[] bytes, DocxStyle style) {
        super(style);
        this.bytes = bytes;
    }
}
