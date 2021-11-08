package wgy.utils;


import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;


public class ExcelUtils {
    /**
     * 获取 Excel 文件表头信息
     *
     * @param fileUrl
     * @return
     * @throws Exception
     */
    public static Set<String> getExcelHeaders(String fileUrl) throws Exception {
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getLastRowNum());
        //获取 excel 第一行数据（表头）
        Row row = sheet.getRow(0);
        //存放表头信息
        Set<String> set = new HashSet<>();
        //算下有多少列
        int colCount = sheet.getRow(0).getLastCellNum();
        System.out.println(colCount);
        for (int j = 0; j < colCount; j++) {
            Cell cell = row.getCell(j);
            String cellValue = cell.getStringCellValue().trim();
            set.add(cellValue);
        }
        return set;
    }

    public static List<String> getExcelHeadersOfList(String fileUrl) throws Exception {
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getLastRowNum());
        //获取 excel 第一行数据（表头）
        Row row = sheet.getRow(0);
        //存放表头信息
        List<String> set = new ArrayList<>();
        //算下有多少列
        int colCount = sheet.getRow(0).getLastCellNum();
        System.out.println(colCount);
        for (int j = 0; j < colCount; j++) {
            Cell cell = row.getCell(j);
            String cellValue = cell.getStringCellValue().trim();
            set.add(cellValue);
            int num = 0;
            for (int i = 0; i < set.size(); i++) {
                if (set.get(i).equals(cellValue)) {
                    num++;
                }
            }
            if (num > 1) {
                set.remove(set.size() - 1);
            }

        }
        return set;
    }

    public static List<String> getExcelHeadersOfListByFile(File file) throws Exception {
//        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getLastRowNum());
        //获取 excel 第一行数据（表头）
        Row row = sheet.getRow(0);
        //存放表头信息
        List<String> set = new ArrayList<>();
        //算下有多少列
        int colCount = sheet.getRow(0).getLastCellNum();
        System.out.println(colCount);
        for (int j = 0; j < colCount; j++) {
            Cell cell = row.getCell(j);
            String cellValue = cell.getStringCellValue().trim();
            set.add(cellValue);
            int num = 0;
            for (int i = 0; i < set.size(); i++) {
                if (set.get(i).equals(cellValue)) {
                    num++;
                }
            }
            if (num > 1) {
                set.remove(set.size() - 1);
            }

        }
        return set;
    }

    /**
     * 获取 Excel 文件信息(除去表头)
     *
     * @param fileUrl
     * @return
     * @throws Exception
     */
    public static List<List<String>> getExcelData(String fileUrl) throws Exception {
        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        //获取 Excel 中 sheet 的行数
        int rowNum = sheet.getLastRowNum();
        List<List<String>> resList = new ArrayList<>();
        //负责标记检测到空行时,跳过
        boolean flag = false;
        for (int i = 1; i <= rowNum; i++) {
            //默认认为此行为空行
            flag = true;
            Row row = sheet.getRow(i);
            //过滤空行
            if (row == null) {
                continue;
            }
            //创建列表，负责装纳一行数据
            List<String> list = new ArrayList<>();
            //获取列数
            int colCount = sheet.getRow(i).getLastCellNum();
            for (int j = 0; j < colCount; j++) {
                //获得制定空格
                Cell cell = row.getCell(j);
                String cellValue = "";
                //如果存在空格内有内容,就将标志位设置为 false，表示这一行不是空行
                if (!(cell == null)) {
                    cellValue = getStringCellValue(cell);
                    if (!"".equals(cellValue)) {
                        flag = false;
                    }
                }
                list.add(cellValue);
            }
            if (!flag) {
                resList.add(list);
            } else {
                continue;
            }
        }
        return resList;
    }

    /**
     * 获取 Excel 文件信息(除去表头)
     *
     * @param file
     * @return
     * @throws Exception
     */
    public static List<List<String>> getExcelDataByFile(File file) throws Exception {
//        File file = new File(fileUrl);
        InputStream is = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);
        //获取 Excel 中 sheet 的行数
        int rowNum = sheet.getLastRowNum();
        List<List<String>> resList = new ArrayList<>();
        //负责标记检测到空行时,跳过
        boolean flag = false;
        for (int i = 1; i <= rowNum; i++) {
            //默认认为此行为空行
            flag = true;
            Row row = sheet.getRow(i);
            //过滤空行
            if (row == null) {
                continue;
            }
            //创建列表，负责装纳一行数据
            List<String> list = new ArrayList<>();
            //获取列数
            int colCount = sheet.getRow(i).getLastCellNum();
            for (int j = 0; j < colCount; j++) {
                //获得制定空格
                Cell cell = row.getCell(j);

                String cellValue = "";

                //如果存在空格内有内容,就将标志位设置为 false，表示这一行不是空行
                if (!(cell == null)) {
                    CellType cellType = cell.getCellType();
                    Object o = null;
                    if (cellType == NUMERIC) {
                        if (DateUtil.isCellDateFormatted(cell)) {
                            o = cell.getDateCellValue();
                        } else {
                            Double val = new Double(cell.getNumericCellValue());
                            // 兼容科学计数法
                            o = String.valueOf(val).indexOf("E") > -1 ? new BigDecimal(val) : val;
                        }
                    }
                    if (o != null) {
                        cellValue = o.toString();
                    } else {
                        cellValue = getStringCellValue(cell);

                    }
                    if (!"".equals(cellValue)) {
                        flag = false;
                    }
                }
                list.add(cellValue);
            }
            if (!flag) {
                resList.add(list);
            } else {
                continue;
            }
        }
        return resList;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    public static String getStringCellValue(Cell cell) {
        String strCell = "";
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                strCell = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue()).trim();
                break;
            case BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        return strCell;
    }

    public static String changeLongDouble(Double v, Integer integer) {
        if (v == null || integer < 0) {
            return "";
        }
        if (integer == null || integer == 0) {
            return new DecimalFormat("0").format(v);
        }
        String num = "0.";
        integer = integer + 2;
        int strLen = num.length();
        if (strLen < integer) {
            while (strLen < integer) {
                StringBuffer sb = new StringBuffer();
//                sb.append("0").append(num);// 左补0
                sb.append(num).append("0");//右补0
                num = sb.toString();
                strLen = num.length();
            }
        }
        return new DecimalFormat(num).format(v);
    }
}
