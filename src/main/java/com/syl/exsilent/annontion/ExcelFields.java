package com.syl.exsilent.annontion;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelFields {
    ExcelField[] value();
}
