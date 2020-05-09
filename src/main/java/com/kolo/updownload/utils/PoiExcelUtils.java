package com.kolo.updownload.utils;

import com.kolo.updownload.vo.BaseCodeVo;
import com.kolo.updownload.vo.BaseDataItemVo;
import com.kolo.updownload.vo.TargetDataItemVo;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class PoiExcelUtils {

    public static <T> XSSFWorkbook createWorkbook(String path, List<T> list, Class<T> clz) throws IOException, InvalidFormatException, IllegalAccessException, IntrospectionException, InvocationTargetException {
        String filePath = PoiExcelUtils.class.getClassLoader().getResource("model" + "/" + path).getPath();
        XSSFWorkbook workbook;
        workbook = (XSSFWorkbook) WorkbookFactory.create(new FileInputStream(new File(filePath)));

        if (clz== BaseDataItemVo.class){
            dataIteam(workbook, (List<BaseDataItemVo>) list, clz);
        }else if (clz == BaseCodeVo.class){
            dataCode(workbook, (List<BaseCodeVo>) list, clz);
        }else if(clz == TargetDataItemVo.class){
            targetDataItem(workbook, (List<TargetDataItemVo>) list, clz);
        }
        return workbook;
    }

    private static void dataIteam(XSSFWorkbook workbook, List<BaseDataItemVo> list, Class clz) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < list.size(); i++) {
            XSSFRow dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
            BaseDataItemVo obj = list.get(i);
            int k = 0;
            Field[] fields = clz.getDeclaredFields();
            for (int j=0;j<fields.length;j++){
                PropertyDescriptor pd = new PropertyDescriptor(fields[k].getName(), clz);
                Method rM = pd.getReadMethod();
                String field = (String)rM.invoke(obj);
                addCell(dataRow, k++,field);
            }

        }
    }

    private static void dataCode(XSSFWorkbook workbook, List<BaseCodeVo> list, Class clz) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < list.size(); i++) {
            XSSFRow dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
            BaseCodeVo obj = list.get(i);
            int k=0;
            Field[] fields = clz.getDeclaredFields();
            for (int j=0;j<fields.length;j++){
                PropertyDescriptor pd = new PropertyDescriptor(fields[k].getName(), clz);
                Method rM = pd.getReadMethod();
                String field = (String)rM.invoke(obj);
                addCell(dataRow, k++,field);
            }

        }
    }

    private static void targetDataItem(XSSFWorkbook workbook, List<TargetDataItemVo> list, Class clz) throws IntrospectionException, InvocationTargetException, IllegalAccessException {
        XSSFSheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < list.size(); i++) {
            XSSFRow dataRow = sheet.createRow(sheet.getLastRowNum() + 1);
            TargetDataItemVo obj = list.get(i);
            int k = 0;
            Field[] fields = clz.getDeclaredFields();
            for (int j=0;j<fields.length;j++){
                PropertyDescriptor pd = new PropertyDescriptor(fields[k].getName(), clz);
                Method rM = pd.getReadMethod();
                String field = (String)rM.invoke(obj);
                addCell(dataRow, k++,field);
            }
        }
    }

    private static void addCell(XSSFRow dataRow, int num, String value){
        XSSFCell cell1 = dataRow.createCell(num);
        cell1.setCellValue(StringUtils.isNotBlank(value)?value:" ");
    }

    public static <T> List<T> importExcel(InputStream inputStream, int sheetNum, int TitleRows, int HeadRows, Class<T> clz) throws IOException, IllegalAccessException, InstantiationException, IntrospectionException, InvocationTargetException {
        List<T> list = new ArrayList<>();
        XSSFWorkbook wb = new XSSFWorkbook(inputStream);
        XSSFSheet sht0 = wb.getSheetAt(sheetNum);
        XSSFRow r = null;
        Field[] fields = clz.getDeclaredFields();
        for(int i=sht0.getFirstRowNum()+TitleRows+HeadRows;i<=sht0.getLastRowNum();i++){
            r = sht0.getRow(i);
            T obj = clz.newInstance();
            for (int j = 0,k=0; j<fields.length; j++,k++){
                PropertyDescriptor pd = new PropertyDescriptor(fields[k].getName(), clz);
                Method rM = pd.getWriteMethod();
                String cellValue;
                if (r.getCell(j)==null){
                    cellValue = "";
                }else {
                    cellValue = getCellValue(r.getCell(j));
                }
                rM.invoke(obj, cellValue);
            }
            list.add(obj);

        }
        return list;
    }

    public static <T> String getCellValue(XSSFCell cell){
        CellType cellType = cell.getCellTypeEnum();
        String cellValue = "";
        switch (cellType) {
            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;

            case FORMULA:
                try {
                    cellValue = cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    cellValue = "#N/A";
                }
                break;
            case ERROR:
                cellValue = "#N/A";
                break;
            default:
                try {
                    cellValue = cell.getStringCellValue();
                }catch (Exception e){
                    e.printStackTrace();
                }

        }
        return cellValue.trim();
    }

}
