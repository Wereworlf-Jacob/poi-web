package com.poi.component.poioperate.utils;

import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName ExcelReader
 * @Description excel文档读取工具类
 * @Author Jacob
 * @Date 2020/4/10 10:10
 * @Version 1.0
 **/
public class ExcelReader {

    //.xls结尾的excel文件
    private static final String XLS = "xls";

    //.xlsx结尾的excel文件
    private static final String XLSX = "xlsx";

    /**
     * 根据文件后缀名类型获取对应的工作簿对象
     * @param inputStream 读取文件的输入流
     * @param fileType 文件后缀名（xls 或 xlsx）
     * @return {@link Workbook} 包含文件数据的工作簿对象
     * @throws IOException
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static Workbook getWorkbook(InputStream inputStream, String fileType) throws IOException {
        Workbook workbook = null;
        if (XLS.equalsIgnoreCase(fileType)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (XLSX.equalsIgnoreCase(fileType)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook;
    }

    /**
     * 根据文件路径读取Excel文件信息
     * @param fileName 文件全路径信息
     * @param clazz 要转化成VO的类信息{VO类信息属性上下顺序要与excel列顺序一致，此值传null则会返回{"列号":"值"}格式的数据}
     * @return {@link List} 返回结果列表
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static <T>List<T> readExcel(String fileName, Class<T> clazz){
        FileInputStream inputStream = null;
        //获取Excel后缀名
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
        //获取Excel文件
        File excelFile = new File(fileName);
        if (!excelFile.exists()) {
            System.out.println("指定的excel文件不存在！@" + fileName);
            return null;
        }
        //获取Excel工作簿
        try {
            inputStream = new FileInputStream(excelFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            try {
                inputStream.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        return readExcel(inputStream, fileType, clazz);
    }

    /**
     * 根据文件流，解析Excel数据
     * @param file
     * @param clazz 要转化成VO的类信息{VO类信息属性上下顺序要与excel列顺序一致，此值传null则会返回{"列号":"值"}格式的数据}
     * @return {@link List}
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static <T>List<T> readExcel(MultipartFile file, Class<T> clazz) {
        //获取后缀名
        String fileName = file.getOriginalFilename();
        if (StringUtils.isEmpty(fileName) || fileName.lastIndexOf(".") < 0) {
            System.out.println("解析Excel失败，因为获取到的Excel文件名非法！ @" + fileName);
            return null;
        }
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
        try {
            return readExcel(file.getInputStream(), fileType, clazz);
        } catch (IOException e) {
            System.out.println("解析Excel失败，因为无法获取到文件的输入流！错误信息：" + e.getMessage());
        }
        return null;
    }

    /**
     * 读取Excel文件内容
     * @param inputStream 包含Excel文件内容的输入流
     * @param fileType 文件后缀名（xls 或 xlsx）
     * @param clazz 要转成VO的类信息
     * @return {@link List} 读取结果列表，读取失败时返回null
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static <T>List<T> readExcel(InputStream inputStream, String fileType, Class<T> clazz){
        Workbook workbook = null;
        try {
            workbook = getWorkbook(inputStream, fileType);
            //读取Excel中的数据
            List resultDateList = parseExcel(workbook, clazz);
            return resultDateList;
        } catch (Exception e) {
            System.out.println("解析Excel失败，错误信息：" + e.getMessage());
            return null;
        } finally {
            try {
                if (null != workbook) {
                    workbook.close();
                }
                if (null != inputStream) {
                    inputStream.close();
                }
            } catch (Exception e) {
                System.out.println("关闭数据流出错！错误信息：" + e.getMessage());
            }
        }
    }

    /**
     * 解析Excel数据
     * @param workbook Excel工作簿对象
     * @param clazz 要转化成VO对象的类信息
     * @return {@link List} 解析结果
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static <T>List<T> parseExcel(Workbook workbook, Class<T> clazz) throws Exception {
        List resultDataList = new ArrayList();
        //解析sheet
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);
            //校验sheet是否合法
            if (sheet == null) {
                continue;
            }
            //获取第一行数据
            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);
            if (null ==firstRow) {
                System.out.println("解析Excel失败，在第一行没有读取到任何数据");
            }

            //解析每一行的数据，构造数据对象
            int rowStart = firstRowNum + 1;
            //获取sheet页不为空的行数
            int rowEnd = sheet.getPhysicalNumberOfRows();
            for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                Object o = convertRowToData(row, clazz);
                if (null == o) {
                    System.out.println("第 " + row.getRowNum() + "行数据不合法，已忽略！");
                    continue;
                }
                resultDataList.add(o);
            }
        }
        return resultDataList;
    }

    /**
     * 提取每一行需要的数据，构造成为一个结果对象
     *
     * 当该行中有单元格的数据为空或不合法时，忽略该行数据
     * @param row 行数据
     * @param clazz 要转换成VO的类信息
     * @return {@link Object} 解析后的行数据对象，行数据错误时返回null
     * @throws Exception
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static <T> T convertRowToData(Row row, Class<T> clazz) throws Exception {
        Cell cell;
        if (clazz == null) { //如果clazz信息是空，那么返回key = 列号，value = 值的json数据
            JSONObject json = new JSONObject();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                json.put(String.valueOf(i), convertCellValueToString(row.getCell(i)));
            }
            return (T) json;
        } else { //否则获取class的set方法信息，然后按照顺序，列号对应vo属性从上到下排列的顺序
            T obj = clazz.newInstance();
            List<Method> methods = ReflectUtil.getClassSetMethod(clazz);
            for (int i = 0; i < methods.size(); i++) {
                methods.get(i).invoke(obj, convertCellValueToString(row.getCell(i)));
            }
            return obj;
        }
    }

    /**
     * 将单元格内容转化为字符串
     * @param cell 
     * @return {@link String}
     * @throws 
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC: //数字
                Double doubleValue = cell.getNumericCellValue();

                //格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING: //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN: //布尔
                Boolean bool = cell.getBooleanCellValue();
                returnValue  = bool.toString();
                break;
            case BLANK: //空值
                break;
            case FORMULA: //公式
                break;
            case ERROR: //出现错误
                break;
            default: //其他默认情况
                break;
        }
        return returnValue;
    }
}
