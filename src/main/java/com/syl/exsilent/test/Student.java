package com.syl.exsilent.test;

import com.syl.exsilent.annontion.ExcelCell;
import com.syl.exsilent.annontion.ExcelField;
import com.syl.exsilent.annontion.ExcelFields;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.Font;

@Data
@Builder
public class Student {
    @ExcelField(index = 0, name = "序号", fontBold = true, fontColor = Font.COLOR_RED, rowCount = 2)
    @ExcelCell(index = 0, font = "华文楷体")
    private String index;

    @ExcelField(index = 1, name = "姓名", fontBold = true, fontColor = Font.COLOR_RED, rowCount = 2)
    @ExcelCell(index = 1, font = "华文楷体")
    private String name;

    @ExcelField(index = 2, name = "年龄", fontBold = true, fontColor = Font.COLOR_RED, rowCount = 2)
    @ExcelCell(index = 2, font = "华文楷体")
    private String age;

    @ExcelField(index = 3, name = "性别", fontBold = true, fontColor = Font.COLOR_RED, rowCount = 2)
    @ExcelCell(index = 3, font = "华文楷体")
    private String gender;

    @ExcelFields({
            @ExcelField(index = 4, name = "成绩", fontBold = true, fontColor = Font.COLOR_RED, colCount = 3),
            @ExcelField(parent = 4, index = 5, name = "语文", fontBold = true, fontColor = Font.COLOR_RED),
            @ExcelField(parent = 4, index = 6, name = "数学", fontBold = true, fontColor = Font.COLOR_RED),
            @ExcelField(parent = 4, index = 7, name = "英语", fontBold = true, fontColor = Font.COLOR_RED)
    })
    @ExcelCell(index = 4, font = "华文楷体")
    private String grade1;
    @ExcelCell(index = 5, font = "华文楷体")
    private String grade2;
    @ExcelCell(index = 6, font = "华文楷体")
    private String grade3;
}
