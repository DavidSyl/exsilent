package com.syl.exsilent.annontion;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * ExcelField
 * 用于数据字段名（表头）的设置
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {
    /**
     * 序号
     */
    int index();

    /**
     * 名称
     */
    String name() default "";

    /**
     * 自动换行
     */
    boolean wrapText() default true;

    /**
     * 水平布局
     */
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;

    /**
     * 垂直布局
     */
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;

    /**
     * 占行数
     */
    int rowCount() default 1;

    /**
     * 占列数
     */
    int colCount() default 1;

    /**
     * 字体
     */
    String font() default "宋体";

    /**
     * 字号
     */
    short fontSize() default 12;

    /**
     * 字体颜色
     */
    short fontColor() default 32767;

    /**
     * 是否加粗
     */
    boolean fontBold() default false;

    /**
     * 是否斜体
     */
    boolean fontItalic() default false;

    /**
     * 父字段
     */
    int parent() default -1;
}
