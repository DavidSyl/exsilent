package com.syl.exsilent.annontion;

import java.lang.annotation.*;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelMeasure {
    /**
     * 行高（单位：磅）
     */
    float[] height() default {};

    /**
     * 列宽（单位：字符）
     */
    int[] width() default {};
}
