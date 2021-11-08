package wgy.action;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import wgy.entity.User;
import wgy.utils.ExcelUtils;

import static org.apache.poi.ss.usermodel.CellType.STRING;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.*;

public class read {
    @Test
    public void get() throws IOException {
        FileInputStream fileInputStream =
                new FileInputStream(new File("D:\\11.xlsx"));
        //2.使用 xssf 创建 workbook
        XSSFWorkbook excel = new XSSFWorkbook(fileInputStream);
        //3.根据索引获取sheet
        XSSFSheet sheet = excel.getSheetAt(0);
        //4.遍历row
        int i = 0;
        for (Row row : sheet) {
            i++;
//            System.out.println(row);
            System.out.println("第" + i + "行");
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    //判断读取的数据中是否有String类型的
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        /*
                        判断是否读取到了日期数据：
                        如果是那就进行格式转换，否则会读取的科学计数值
                        不是就输出number 数字
                         */
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            Date date = (Date) cell.getDateCellValue();
                            //格式转换
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            String format = sdf.format(date);
                            System.out.println(format);
                        } else {
                            System.out.println(cell.getNumericCellValue());
                        }
                        break;
                }
            }
        }
        excel.close();

    }

    //    public static void main(String[] args) throws Exception {
//        List<List<String>> excelData = ExcelUtils.getExcelData("D:\\11.xlsx");
//        List<HashMap<String,String>> result = new ArrayList<>();
//        List<User> userList =new ArrayList<>();
//        List<String> excelHeadersOfList = ExcelUtils.getExcelHeadersOfList("D:\\11.xlsx");
//        for (int i = 0; i < excelData.size(); i++) {
//            HashMap<String,String> map = new HashMap<>();
//            for (int j = 0; j < excelHeadersOfList.size(); j++) {
//                map.put(excelHeadersOfList.get(j),excelData.get(i).get(j));
//            }
//            result.add(map);
//        }
//
//        for (int j = 0; j < result.size(); j++) {
//            User user =new User();
//            for (int i = 0; i < excelHeadersOfList.size(); i++) {
//                Field field = user.getClass().getDeclaredField(String.valueOf(excelHeadersOfList.get(i)));
//                field.setAccessible(true);
//                String s = result.get(j).get(excelHeadersOfList.get(i));
//                field.set(user, s);
//            }
//            userList.add(user);
//        }
//        System.out.println("111");
//
//    }
    //根据excel获取数据
    public List<User> getDate() throws Exception {
        List<List<String>> excelData = ExcelUtils.getExcelData("D:\\11.xlsx");
        List<HashMap<String, String>> result = new ArrayList<>();
        List<User> userList = new ArrayList<>();
        List<String> excelHeadersOfList = ExcelUtils.getExcelHeadersOfList("D:\\11.xlsx");
        for (int i = 0; i < excelData.size(); i++) {
            HashMap<String, String> map = new HashMap<>();
            for (int j = 0; j < excelHeadersOfList.size(); j++) {
                map.put(excelHeadersOfList.get(j), excelData.get(i).get(j));
            }
            result.add(map);
        }

        for (int j = 0; j < result.size(); j++) {
            User user = new User();
            for (int i = 0; i < excelHeadersOfList.size(); i++) {
                Field field = user.getClass().getDeclaredField(String.valueOf(excelHeadersOfList.get(i)));
                field.setAccessible(true);
                String s = result.get(j).get(excelHeadersOfList.get(i));
                field.set(user, s);
            }
            userList.add(user);
        }
        return userList;
    }

    public List<User> getDateByFile(File file) throws Exception {
        List<List<String>> excelData = ExcelUtils.getExcelDataByFile(file);
        List<HashMap<String, String>> result = new ArrayList<>();
        List<User> userList = new ArrayList<>();
        List<String> excelHeadersOfList = ExcelUtils.getExcelHeadersOfListByFile(file);
        for (int i = 0; i < excelData.size(); i++) {
            HashMap<String, String> map = new HashMap<>();
            for (int j = 0; j < excelHeadersOfList.size(); j++) {
                map.put(excelHeadersOfList.get(j), excelData.get(i).get(j));
            }
            result.add(map);
        }

        for (int j = 0; j < result.size(); j++) {
            User user = new User();
            for (int i = 0; i < excelHeadersOfList.size(); i++) {
                Field field = user.getClass().getDeclaredField(String.valueOf(excelHeadersOfList.get(i)));
                field.setAccessible(true);
                String s = result.get(j).get(excelHeadersOfList.get(i));
                field.set(user, s);
            }
            userList.add(user);
        }
        return userList;
    }

}
