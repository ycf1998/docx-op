package com.money.docx.item;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.Setter;

import java.util.List;

/**
 * @author : money
 * @version : 1.0.0
 * @description : 多克斯表
 * @createTime : 2022-08-20 15:50:32
 */
@Getter
@Setter
@RequiredArgsConstructor
public class DocxTable extends DocxItem {

    /**
     * 行
     */
    private final List<DocxRow> rows;

    /**
     * 获取行数
     *
     * @return int
     */
    public int getRowCount() {
        return rows.size();
    }

    /**
     * 获得列数
     *
     * @return int
     */
    public int getColumnCount() {
        return rows.stream().map(row -> row.getCells().size()).findFirst().orElse(0);
    }
}
