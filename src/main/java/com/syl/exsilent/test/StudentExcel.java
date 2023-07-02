package com.syl.exsilent.test;

import com.syl.exsilent.annontion.ExcelMeasure;
import com.syl.exsilent.annontion.ExcelModule;
import lombok.Builder;
import lombok.Data;

import java.util.List;

@ExcelMeasure(width = {8, 10}, height = {30F, -1})
@Data
@Builder
public class StudentExcel {
    @ExcelModule(index = 0, rowCount = 2, colCount = 7, fontBold = true, fontItalic = true, fontSize = 14, font = "黑体")
    private String title;
    @ExcelModule(index = 1)
    private List<Student> students;
}
