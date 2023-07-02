package com.syl.exsilent.annontion;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * ExcelModule
 * 用于组合多个数据集合模块
 * 仅支持String、List<T>类型的字段。
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelModule {
    /**
     * 序号
     */
    int index();
    /**
     * 单元格高度
     */
    double height() default 15.75;
    /**
     * 单元格宽度
     */
    double wight() default 8.44;
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
    short fontColor() default Font.COLOR_NORMAL;
    /**
     * 是否加粗
     */
    boolean fontBold() default false;
    /**
     * 是否斜体
     */
    boolean fontItalic() default false;
}
