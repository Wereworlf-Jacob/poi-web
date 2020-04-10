package com.poi.component.poioperate.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xwpf.usermodel.Borders;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName ExcelWriter
 * @Description 生成Excel并写入数据
 * @Author Jacob
 * @Version 1.0
 * @since 2020/4/10 14:23
 **/
public class ExcelWriter {

    private static List<String> CELL_HEADS; //列头

    static {
        //类装载时就载入指定好的列头信息，如有需要，可以考虑做成动态生成的列头
        CELL_HEADS = new ArrayList<>();
        CELL_HEADS.add("1");
        CELL_HEADS.add("2");
        CELL_HEADS.add("3");
        CELL_HEADS.add("4");
        CELL_HEADS.add("5");
    }

    /**
     * 生成Excel并写入数据信息
     * @param dataList 数据列表
     * @return {@link Workbook} 写入数据后的工作簿对象
     * @throws 
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static Workbook exportData(List<?> dataList){
        //生成xlsx的Excel
        Workbook workbook = new SXSSFWorkbook();

        //如需生成xls的Excel,请使用下面的工作簿对象，注意后续输出时文件后缀名也需更改为.xls
        //Workbook workbook = new HSSFWorkbook();

        //生成sheet表，写入第一行的列头
        Sheet sheet = buildDataSheet(workbook);

        try {
            //构建每行的数据内容
            int rowNum = 1;
            for (Object o : dataList) {
                if (o == null) {
                    continue;
                }
                //输出行数据
                Row row = sheet.createRow(rowNum++);
                convertDataToRow(o, row);
            }
        } catch (Exception e) {
            System.out.println("写入Excel数据失败，错误信息：" + e.getMessage());
        }

        return workbook;
    }


    /**
     * 生成sheet表，并写入第一行数据（列头）
     * @param workbook 工作簿对象
     * @return {@link Sheet} 已经写入列头的sheet
     * @throws 
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static Sheet buildDataSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet();
        //设置列头宽度
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            sheet.setColumnWidth(i, 4000);
        }
        //设置默认行高
        sheet.setDefaultRowHeight((short) 400);
        //构建头单元格样式
        CellStyle cellStyle = buildHeadCellStyle(sheet.getWorkbook());
        //写入第一行各列的数据
        Row head = sheet.createRow(0);
        for (int i = 0; i < CELL_HEADS.size(); i++) {
            Cell cell = head.createCell(i);
            cell.setCellValue(CELL_HEADS.get(i));
            cell.setCellStyle(cellStyle);
        }
        return sheet;
    }

    /**
     * 设置第一行列头的样式
     * @param workbook 工作簿对象
     * @return {@link CellStyle} 单元格样式对象
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static CellStyle buildHeadCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        //对齐方式设置
        style.setAlignment(HorizontalAlignment.CENTER);
        //边框颜色和宽度设置
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());//下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); //左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); //有边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());//上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 将数据转化成行
     * @param obj 对象值
     * @param row excel行数据
     * @return 
     * @throws 
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static void convertDataToRow(Object obj, Row row) throws Exception {
        Cell cell;
        Class clazz = obj.getClass();
        List<Method> methods = ReflectUtil.getClassGetMethod(clazz);
        for (int i = 0; i < methods.size(); i++) {
            cell = row.createCell(i);
            Object value = methods.get(i).invoke(obj, null);
            cell.setCellValue(null == value ? "" : value.toString());
        }
    }

}
