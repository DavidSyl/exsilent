package com.syl.exsilent.test;

import cn.hutool.core.collection.ListUtil;
import com.syl.exsilent.builder.ExcelBuilder;
import com.syl.exsilent.builder.SheetBuilder;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * todo 设置边框
 */
public class ExcelTest {

    @Test
    public void test() {

        Student student1 = Student.builder().index("1").name("张三").age("16").gender("男").grade1("75").grade2("85").grade3("89").build();
        Student student2 = Student.builder().index("2").name("李四").age("17").gender("男").grade1("88").grade2("95").grade3("87").build();
        Student student3 = Student.builder().index("3").name("王五").age("15").gender("男").grade1("90").grade2("79").grade3("85").build();
        Student student4 = Student.builder().index("4").name("小美").age("16").gender("女").grade1("79").grade2("81").grade3("92").build();

        List<Student> students = new ArrayList<>();
        students.add(student1);
        students.add(student2);
        students.add(student3);
        students.add(student4);

        StudentExcel studentExcel1 = StudentExcel.builder().title("一年级1班学生信息").students(students).build();
        StudentExcel studentExcel2 = StudentExcel.builder().title("一年级2班学生信息").students(students).build();
        StudentExcel studentExcel3 = StudentExcel.builder().title("二年级1班学生信息").students(students).build();
        StudentExcel studentExcel4 = StudentExcel.builder().title("二年级2班学生信息").students(students).build();

        // 1.输出多个sheet
        ExcelBuilder excelBuilder = new ExcelBuilder(ExcelBuilder.ExcelType.XLS);
        excelBuilder.sheetBuilder("一年级学生信息").append(studentExcel1).append(studentExcel2).append(students);
        excelBuilder.sheetBuilder("二年级学生信息").append(studentExcel3);
        excelBuilder.writeAndClose(new File("C:\\Users\\WeiJiahao\\Desktop\\1.xls"));

        // 2.增加数据后重新构建
        excelBuilder.sheetBuilder("二年级学生信息").append(studentExcel4);
        excelBuilder.writeAndClose("C:\\Users\\WeiJiahao\\Desktop\\2.xls");

        // 3.获取sheet进行特殊处理，例：通过setZoom设置缩放比例
        excelBuilder.hssfSheet("一年级学生信息").setZoom(1, 2);
        excelBuilder.writeAndClose("C:\\Users\\WeiJiahao\\Desktop\\3.xls");

        // 4.获取workbook进行特殊处理，
        excelBuilder.workbook().setActiveSheet(1);
        excelBuilder.writeAndClose("C:\\Users\\WeiJiahao\\Desktop\\4.xls");

        // 5.输出xlsx类型Excel文件
        ExcelBuilder excelBuilder1 = new ExcelBuilder(ExcelBuilder.ExcelType.XLSX);
        excelBuilder1.sheetBuilder("一年级学生信息").append(studentExcel1).append(studentExcel2).append(students);
        excelBuilder1.writeAndClose(new File("C:\\Users\\WeiJiahao\\Desktop\\5.xlsx"));
    }

    @Test
    public void AsyncTest() {

        Student student1 = Student.builder().index("1").name("张三").age("16").gender("男").grade1("75").grade2("85").grade3("89").build();
        Student student2 = Student.builder().index("2").name("李四").age("17").gender("男").grade1("88").grade2("95").grade3("87").build();
        Student student3 = Student.builder().index("3").name("王五").age("15").gender("男").grade1("90").grade2("79").grade3("85").build();
        Student student4 = Student.builder().index("4").name("小美").age("16").gender("女").grade1("79").grade2("81").grade3("92").build();

        List<Student> students = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            students.add(student1);
            students.add(student2);
            students.add(student3);
            students.add(student4);
        }

        // 1.输出多个sheet

        int core = Runtime.getRuntime().availableProcessors();
        System.out.println(core);

        ThreadPoolExecutor executor = new ThreadPoolExecutor(4, 8, 0, TimeUnit.SECONDS, new ArrayBlockingQueue<>(10));
        ExcelBuilder excelBuilder = new ExcelBuilder(ExcelBuilder.ExcelType.XLS);
        SheetBuilder sheetBuilder = excelBuilder.sheetBuilder("一年级学生信息");

        List<List<Student>> partition = ListUtil.partition(students, 10);
        for (List<Student> studentList : partition) {
            executor.execute(() -> sheetBuilder.append(studentList));
        }

        executor.shutdown();
//        CountDownLatch countDownLatch = new CountDownLatch(4);
//        try {
//            countDownLatch.await();
//        } catch (InterruptedException e) {
//            e.printStackTrace();
//        }
        while (!executor.isTerminated()) {

        }
        excelBuilder.writeAndClose(new File("C:\\Users\\WeiJiahao\\Desktop\\1-async.xls"));


    }

}
