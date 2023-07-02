package com.syl.exsilent.builder;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * workbook builder
 *
 * @since 2023.6.23
 */
@Slf4j
public class ExcelBuilder {
    /**
     * workbook实例
     */
    private volatile Workbook workbook;
    /**
     * 工作表builder Map
     * k-sheetBuilder name
     * v-sheetBuilder builder
     */
    private final ConcurrentHashMap<String, SheetBuilder> sheetBuilderMap = new ConcurrentHashMap<>();

    /**
     * 构造方法
     *
     * @param excelType excel类型
     */
    public ExcelBuilder(ExcelType excelType) {
        if (excelType == ExcelType.XLS) {
            workbook = new HSSFWorkbook();
        } else if (excelType == ExcelType.XLSX) {
            workbook = new XSSFWorkbook();
        }
    }

    /**
     * 新增sheetBuilder或获取sheetBuilder
     *
     * @param sheetName 工作表名称
     * @return 新的或已有的sheet builder对象
     */
    public SheetBuilder sheetBuilder(String sheetName) {
        if (sheetBuilderMap.containsKey(sheetName)) {
            return sheetBuilderMap.get(sheetName);
        } else {
            synchronized (this.sheetBuilderMap) {
                if (sheetBuilderMap.containsKey(sheetName)) {
                    return sheetBuilderMap.get(sheetName);
                }
                SheetBuilder sheetBuilder = new SheetBuilder(this.workbook, sheetName);
                sheetBuilderMap.put(sheetName, sheetBuilder);
                return sheetBuilder;
            }
        }
    }

    /**
     * 获取sheet
     * 返回前进行构建
     *
     * @param sheetName 工作表名称
     * @return sheet对象
     */
    public Sheet sheet(String sheetName) {
        SheetBuilder sheetBuilder = sheetBuilderMap.get(sheetName);
        if (sheetBuilder != null) {
            sheetBuilder.build();
        }
        return this.workbook.getSheet(sheetName);
    }

    @SneakyThrows
    public HSSFSheet hssfSheet(String sheetName) {
        Sheet sheet = this.sheet(sheetName);
        if (sheet instanceof HSSFSheet) {
            return (HSSFSheet) sheet;
        } else {
            throw new Exception(sheetName + "is not a instance of HSSFSheet");
        }
    }

    @SneakyThrows
    public XSSFSheet xssfSheet(String sheetName) {
        Sheet sheet = this.sheet(sheetName);
        if (sheet instanceof XSSFSheet) {
            return (XSSFSheet) sheet;
        } else {
            throw new Exception(sheetName + "is not a instance of XSSFSheet");
        }
    }

    /**
     * 构造workbook
     */
    public synchronized ExcelBuilder build() {
        this.sheetBuilderMap.values().forEach(SheetBuilder::build);
        return this;
    }

    /**
     * 将workbook内容写入File实例中并关闭workbook
     *
     * @param newFile File实例
     */
    public void writeAndClose(File newFile) {
        this.build();
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(newFile);
        } catch (FileNotFoundException e) {
            log.error(e.getMessage(), e);
        }
        this.writeAndClose(fileOutputStream);
    }

    /**
     * 将workbook内容写入指定路径文件中并关闭workbook
     *
     * @param newFileDir 文件路径
     */
    public void writeAndClose(String newFileDir) {
        this.build();
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(newFileDir);
        } catch (FileNotFoundException e) {
            log.error(e.getMessage(), e);
        }
        this.writeAndClose(fileOutputStream);
    }

    /**
     * 将workbook内容写入输出流中并关闭workbook
     *
     * @param outputStream 输出流
     */
    public void writeAndClose(OutputStream outputStream) {
        this.build();
        try {
            this.workbook.write(outputStream);
        } catch (IOException e) {
            log.error("workbook write failed", e);
        } finally {
            try {
                outputStream.close();
                this.workbook.close();
            } catch (IOException e) {
                log.error("workbook close failed", e);
            }
        }
    }

    /**
     * 获取构造的workbook
     *
     * @return workbook对象
     */
    public Workbook workbook() {
        return this.build().workbook;
    }

    @SneakyThrows
    public HSSFWorkbook hssfWorkbook() {
        Workbook workbook = this.workbook();
        if (workbook instanceof HSSFWorkbook) {
            return (HSSFWorkbook) workbook;
        } else {
            throw new Exception("this workbook is not a instance of HSSFWorkbook");
        }
    }

    @SneakyThrows
    public XSSFWorkbook xssfWorkbook() {
        Workbook workbook = this.workbook();
        if (workbook instanceof XSSFWorkbook) {
            return (XSSFWorkbook) workbook;
        } else {
            throw new Exception("this workbook is not a instance of XSSFWorkbook");
        }
    }

    public enum ExcelType {
        XLS,
        XLSX
    }
}
