package com.money.docx.constant;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 表格合并枚举
 * @createTime : 2022-08-27 12:25:50
 */
@Getter
@RequiredArgsConstructor
public enum TbMergeEnum {

    /**
     * 开始
     */
    RESTART("restart"),
    /**
     * 继续
     */
    CONTINUE("continue");

    private final String key;
}
