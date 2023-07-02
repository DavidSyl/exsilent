package com.syl.exsilent.model;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 单元格数据模型
 *
 * @since 2023.6.23
 */
@Data
public class CellModel {
    /**
     * 数据
     */
    private String value;
    /**
     * 样式
     */
    private CellStyle style;
    /**
     * 行索引
     */
    private int rowNum;
    /**
     * 列索引
     */
    private int colNum;
}
