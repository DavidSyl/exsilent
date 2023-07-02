package com.syl.exsilent.builder;

import com.syl.exsilent.annontion.*;
import com.syl.exsilent.model.CellModel;
import com.syl.exsilent.util.ExcelUtil;
import lombok.Builder;
import lombok.Data;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * 工作表 builder
 *
 * @since 2023.6.23
 */
@Slf4j
public class SheetBuilder {
    /**
     * 工作表实例引用
     */
    private final Sheet sheet;
    /**
     * workbook实例引用
     */
    private final Workbook workbook;
    /**
     * 单元格列表-存储构造过程中生成的单元格
     */
    private final List<CellModel> cells = new ArrayList<>();
    /**
     * 合并区域列表-存储构造过程中生成的单元格合并数据
     */
    private final List<CellRangeAddress> regions = new ArrayList<>();
    /**
     * 全局列数
     */
    private final AtomicInteger colNum = new AtomicInteger(0);
    /**
     * 全局待插入行数map
     * k-colNum
     * v-rowNum
     */
    private final ConcurrentHashMap<Integer, Integer> colRowMap = new ConcurrentHashMap<>();
    /**
     * 数据是否修改
     */
    private final AtomicBoolean isDataUpdated = new AtomicBoolean(false);
    /**
     * 列宽、行高数据
     */
    private ExcelMeasure measure = null;

    SheetBuilder(Workbook workbook, String sheetName) {
        this.workbook = workbook;
        this.sheet = workbook.createSheet(sheetName);
    }

    /**
     * 向工作表中新增数据
     * 最底层方法
     *
     * @param val   单元格数据
     * @param param 单元格参数
     */
    private synchronized void append(@NonNull String val, @NonNull CommonParam param) {
        this.isDataUpdated.set(true);
        // 样式
        CellStyle style = this.workbook.createCellStyle();
        style.setAlignment(param.getHorizontalAlignment());
        style.setVerticalAlignment(param.getVerticalAlignment());
        style.setWrapText(param.isWrapText());
        Font font = this.workbook.createFont();
        font.setFontName(param.getFont());
        font.setFontHeightInPoints(param.getFontSize());
        font.setColor(param.getFontColor());
        font.setBold(param.isFontBold());
        font.setItalic(param.isFontItalic());
        style.setFont(font);

        // 单元格合并
        int rowCount = param.getRowCount();
        int colCount = param.getColCount();
        int curCol = this.colNum.get();
        int rowNum = this.getRowNum(curCol);
        if (rowCount != 1 || colCount != 1) {
            CellRangeAddress address = new CellRangeAddress(rowNum, rowNum + rowCount - 1, curCol, curCol + colCount - 1);
            this.regions.add(address);
        }

        CellModel cellModel = new CellModel();
        cellModel.setValue(val);
        cellModel.setStyle(style);
        cellModel.setRowNum(rowNum);
        cellModel.setColNum(curCol);
        this.cells.add(cellModel);
    }

    /**
     * 向工作表中新增表头数据
     *
     * @param excelField 表头注解参数
     */
    private void append(@NonNull ExcelField excelField) {
        this.append(excelField.name(), CommonParam.fromXlsField(excelField));
    }

    /**
     * 向工作表中新增嵌套表头数据
     *
     * @param excelFields 嵌套表头注解参数
     */
    private void append(@NonNull ExcelFields excelFields) {
        ExcelField[] excelFieldArr = excelFields.value();
        if (excelFieldArr.length == 0) {
            return;
        }
        Arrays.sort(excelFieldArr, Comparator.comparingInt(ExcelField::parent));
        if (excelFieldArr[0].parent() != -1) {
            log.warn("hierarchical fields without root");
            return;
        }
        Map<Integer, List<ExcelField>> tierMap = new TreeMap<>();
        Map<Integer, Integer> indexTierMap = new HashMap<>();
        for (ExcelField excelField : excelFieldArr) {
            if (excelField.parent() == -1) {
                indexTierMap.put(excelField.index(), 0);
            } else {
                Integer parentTier = indexTierMap.get(excelField.parent());
                if (parentTier == null) {
                    log.warn("hierarchical fields index set illegally,{} has not parent", excelField.name());
                } else {
                    indexTierMap.put(excelField.index(), parentTier + 1);
                }
            }
        }

        for (ExcelField excelField : excelFieldArr) {
            Integer tier = indexTierMap.get(excelField.index());
            if (tier != null) {
                tierMap.putIfAbsent(tier, new ArrayList<>());
                tierMap.get(tier).add(excelField);
            }
        }

        tierMap.forEach((k, v) -> {
            int localColNum = this.colNum.get();
            for (ExcelField field : v) {
                this.append(field);
                this.updateRowNum(this.colNum.get(), field.colCount(), field.rowCount());
                this.colNum.set(this.colNum.addAndGet(field.colCount()));
            }
            this.colNum.set(localColNum);
        });
    }

    /**
     * 向工作表中新增列表类型数据
     *
     * @param list 工作表列表数据
     * @param <T>  泛型类
     * @return 当前builder对象
     */
    public <T> SheetBuilder append(@NonNull List<T> list) {
        if (list.isEmpty()) {
            return this;
        }
        Class<?> clazz = list.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();

        // 新增表头数据。按序号排序，过滤未注解字段
        Map<Integer, ExcelFields> xlsFieldsMap = new HashMap<>();
        List<ExcelField> excelFields = Arrays.stream(fields)
                .map(f -> {
                    ExcelField excelField = f.getAnnotation(ExcelField.class);
                    ExcelFields excelFieldsAno = f.getAnnotation(ExcelFields.class);
                    if (excelFieldsAno == null && excelField == null) {
                        return null;
                    }
                    if (excelFieldsAno != null && excelField != null) {
                        log.warn("@excelField and @excelFields can not be used together");
                    }
                    if (excelFieldsAno != null) {
                        ExcelField[] value = excelFieldsAno.value();
                        Optional<ExcelField> minOne = Arrays.stream(value).min(Comparator.comparing(ExcelField::parent));
                        minOne.ifPresent(m -> xlsFieldsMap.put(m.index(), excelFieldsAno));
                        return minOne.orElse(null);
                    }
                    return excelField;
                })
                .filter(Objects::nonNull)
                .sorted(Comparator.comparing(ExcelField::index))
                .toList();

        this.colNum.set(0);
        for (ExcelField excelField : excelFields) {
            int index = excelField.index();
            if (xlsFieldsMap.containsKey(index)) {
                // xlsFields处理
                this.append(xlsFieldsMap.get(index));
            } else {
                this.append(excelField);
                this.updateRowNum(this.colNum.get(), excelField.colCount(), excelField.rowCount());
                this.colNum.set(this.colNum.addAndGet(excelField.colCount()));
            }
        }

        // 新增内容数据
        List<CommonParam> params = Arrays.stream(fields)
                .map(f -> {
                    ExcelCell excelCell = f.getAnnotation(ExcelCell.class);
                    if (null == excelCell) {
                        return null;
                    }
                    CommonParam param = CommonParam.fromXlsCell(excelCell);
                    param.setFieldName(f.getName());
                    return param;
                })
                .filter(Objects::nonNull)
                .sorted(Comparator.comparing(CommonParam::getIndex))
                .toList();

        for (T t : list) {
            this.colNum.set(0);
            for (CommonParam param : params) {
                String val = (String) ExcelUtil.getFieldVal(t, param.getFieldName());
                this.append(val, param);
                this.updateRowNum(this.colNum.get(), param.getColCount(), param.getRowCount());
                this.colNum.set(this.colNum.addAndGet(param.getColCount()));
            }
        }

        // 设置列宽、行高
        if (this.measure == null) {
            ExcelMeasure excelMeasure = clazz.getAnnotation(ExcelMeasure.class);
            if (excelMeasure != null) {
                this.measure = excelMeasure;
            }
        }

        return this;
    }

    /**
     * 向工作表中新增对象模型数据
     *
     * @param model 对象模型数据-可以包含多个表格模块
     * @param <T>   泛型类
     * @return 当前builder对象
     */
    public <T> SheetBuilder append(@NonNull T model) {
        Class<?> clazz = model.getClass();
        Field[] fields = clazz.getDeclaredFields();

        // 对字段进行过滤排序，转换为参数列表
        List<CommonParam> params = Arrays.stream(fields)
                .map(f -> {
                    ExcelModule excelModule = f.getAnnotation(ExcelModule.class);
                    if (null == excelModule) {
                        return null;
                    }
                    CommonParam param = CommonParam.fromXlsModule(excelModule);
                    param.setFieldName(f.getName());
                    return param;
                })
                .filter(Objects::nonNull)
                .sorted(Comparator.comparing(CommonParam::getIndex))
                .toList();

        for (CommonParam param : params) {
            this.colNum.set(0);
            Object val = ExcelUtil.getFieldVal(model, param.getFieldName());
            if (val instanceof List<?>) {
                // 按列表数据处理
                this.append((List<?>) val);
            } else if (val instanceof String) {
                this.append((String) val, param);
                this.updateRowNum(this.colNum.get(), param.getColCount(), param.getRowCount());
            } else {
                log.warn("仅支持List和String类型的字段，无法处理{}字段", param.getFieldName());
            }
        }

        // 设置列宽、行高
        if (this.measure == null) {
            ExcelMeasure excelMeasure = clazz.getAnnotation(ExcelMeasure.class);
            if (excelMeasure != null) {
                this.measure = excelMeasure;
            }
        }

        return this;
    }

    /**
     * 构造工作表
     */
    public synchronized void build() {
        if (!this.isDataUpdated.get()) {
            return;
        }
        Map<Integer, List<CellModel>> map = this.cells.stream()
                .collect(Collectors.groupingBy(CellModel::getRowNum));
        for (Map.Entry<Integer, List<CellModel>> entry : map.entrySet()) {
            Row row = this.sheet.createRow(entry.getKey());
            for (CellModel cellModel : entry.getValue()) {
                Cell cell = row.createCell(cellModel.getColNum(), CellType.STRING);
                cell.setCellValue(cellModel.getValue());
                cell.setCellStyle(cellModel.getStyle());
            }
        }
        this.cells.clear();
        this.setHightAndWidth(this.measure);
        this.regions.forEach(this.sheet::addMergedRegion);
        this.regions.clear();
        this.isDataUpdated.set(false);
    }

    /**
     * 更新 列数-待插入数据行数map
     *
     * @param currentColNum 当前全局列数
     * @param colCount      影响列数
     * @param rowCount      行数变化值
     */
    private void updateRowNum(int currentColNum, int colCount, int rowCount) {
        for (int i = currentColNum; i < currentColNum + colCount; i++) {
            this.colRowMap.putIfAbsent(i, 0);
            this.colRowMap.computeIfPresent(i, (k, v) -> v + rowCount);
        }
    }

    /**
     * 根据列数获取该列待插入数据的行数
     *
     * @param colNum 列数
     * @return 待插入行数
     */
    private int getRowNum(int colNum) {
        if (this.colRowMap.containsKey(colNum)) {
            return this.colRowMap.get(colNum);
        } else {
            this.colRowMap.put(colNum, 0);
            return 0;
        }
    }

    /**
     * 设置行高和列宽
     *
     * @param excelMeasure 参数
     */
    private synchronized void setHightAndWidth(ExcelMeasure excelMeasure) {
        if (excelMeasure == null) {
            return;
        }
        float[] heights = excelMeasure.height();
        int[] widths = excelMeasure.width();
        for (int i = 0, heightsLength = heights.length; i < heightsLength; i++) {
            float height = heights[i];
            Row row = this.sheet.getRow(i);
            if (row == null) {
                row = this.sheet.createRow(i);
            }
            row.setHeightInPoints(height);
        }
        for (int i = 0, widthsLength = widths.length; i < widthsLength; i++) {
            int width = widths[i];
            this.sheet.setColumnWidth(i, width * 256);
        }
    }


    /**
     * 公共参数
     */
    @Data
    @Builder
    static class CommonParam {
        /**
         * 字段名称
         */
        private String fieldName;
        /**
         * 序号
         */
        private int index;
        /**
         * 单元格高度
         */
        private double height;
        /**
         * 单元格宽度
         */
        private double wight;
        /**
         * 自动换行
         */
        private boolean wrapText;
        /**
         * 水平布局
         */
        private HorizontalAlignment horizontalAlignment;
        /**
         * 垂直布局
         */
        private VerticalAlignment verticalAlignment;
        /**
         * 占行数
         */
        private int rowCount;
        /**
         * 占列数
         */
        private int colCount;
        /**
         * 字体
         */
        private String font;
        /**
         * 字号
         */
        private short fontSize;
        /**
         * 字体颜色
         */
        private short fontColor;
        /**
         * 是否加粗
         */
        private boolean fontBold;
        /**
         * 是否斜体
         */
        private boolean fontItalic;

        /**
         * XlsModule转换为通用参数对象
         *
         * @param excelModule XlsModule对象
         * @return 通用参数对象
         */
        public static CommonParam fromXlsModule(@NonNull ExcelModule excelModule) {
            return CommonParam.builder()
                    .index(excelModule.index())
//                    .height(excelModule.height())
//                    .wight(excelModule.wight())
                    .wrapText(excelModule.wrapText())
                    .horizontalAlignment(excelModule.horizontalAlignment())
                    .verticalAlignment(excelModule.verticalAlignment())
                    .rowCount(excelModule.rowCount()).colCount(excelModule.colCount())
                    .font(excelModule.font()).fontSize(excelModule.fontSize())
                    .fontColor(excelModule.fontColor()).fontBold(excelModule.fontBold())
                    .fontItalic(excelModule.fontItalic())
                    .build();
        }

        /**
         * XlsCell转换为通用参数对象
         *
         * @param excelCell XlsCell对象
         * @return 通用参数对象
         */
        public static CommonParam fromXlsCell(@NonNull ExcelCell excelCell) {
            return CommonParam.builder().index(excelCell.index())
//                    .height(excelCell.height())
//                    .wight(excelCell.wight())
                    .wrapText(excelCell.wrapText())
                    .horizontalAlignment(excelCell.horizontalAlignment())
                    .verticalAlignment(excelCell.verticalAlignment())
                    .rowCount(excelCell.rowCount())
                    .colCount(excelCell.colCount())
                    .font(excelCell.font())
                    .fontSize(excelCell.fontSize())
                    .fontColor(excelCell.fontColor())
                    .fontBold(excelCell.fontBold())
                    .fontItalic(excelCell.fontItalic())
                    .build();
        }

        /**
         * XlsField转换为通用参数对象
         *
         * @param excelField XlsField对象
         * @return 通用参数对象
         */
        public static CommonParam fromXlsField(@NonNull ExcelField excelField) {
            return CommonParam.builder().index(excelField.index())
//                    .height(excelField.height())
//                    .wight(excelField.wight())
                    .wrapText(excelField.wrapText())
                    .horizontalAlignment(excelField.horizontalAlignment())
                    .verticalAlignment(excelField.verticalAlignment())
                    .rowCount(excelField.rowCount())
                    .colCount(excelField.colCount())
                    .font(excelField.font())
                    .fontSize(excelField.fontSize())
                    .fontColor(excelField.fontColor())
                    .fontBold(excelField.fontBold())
                    .fontItalic(excelField.fontItalic())
                    .build();
        }
    }
}
