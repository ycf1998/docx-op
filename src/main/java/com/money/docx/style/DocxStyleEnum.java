package com.money.docx.style;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.docx4j.wml.JcEnumeration;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯风格枚举
 * @createTime : 2022-08-20 16:11:40
 */
@Getter
@AllArgsConstructor
public enum DocxStyleEnum {

    DEFAULT_TOC_TITLE(DocxStyle.builder().align(JcEnumeration.CENTER).bold(true).fontSize(20L).build());


    private final DocxStyle docxStyle;

}
