package com.huawei.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Properties;
import java.util.logging.Logger;

public class Main {

    private static Logger logger = Logger.getLogger(Main.class.getName()); // 日志打印类

    public static void main(String[] args) {
        boolean success = false;
        if (args.length != 0){
            String excelFilePath = args[0];
            success = parseExcel(excelFilePath);
        }else{
            logger.warning("请输入待解析的excel文件路径");
        }
        if (success){
            logger.info("写入成功");
        }
    }

    private static boolean parseExcel(String filePath){
        File excelFile = new File(filePath);
        if (!excelFile.exists()) {
            logger.warning("指定的Excel文件不存在！");
            return false;
        }
        InputStream excelInputStream = null;
        InputStream propInputStream = null;
        OutputStream propOutputStream = null;
        try {
            // 如果没有generate.properties则创建一个新的，如果存在则删除再创建新的
            File propFile = new File(".\\generate.properties");
            if (propFile.exists()){
                if(!propFile.delete()){
                    logger.warning("删除properties文件失败");
                }
            }
            if(!propFile.createNewFile()){
                logger.warning("创建properties文件失败");
            }
            Properties prop = new Properties();
            propInputStream = new FileInputStream(propFile);
            propOutputStream = new FileOutputStream(propFile);
            prop.load(propInputStream);

            Workbook wb = null;
            excelInputStream = new FileInputStream(excelFile);
            // 检查excel文件格式，xls或者xlsx
            String excelFileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
            if (excelFileType.equalsIgnoreCase("XLS")) {
                wb = new HSSFWorkbook(excelInputStream);
            } else if (excelFileType.equalsIgnoreCase("XLSX")) {
                wb = new XSSFWorkbook(excelInputStream);
            }
            if (wb == null){
                logger.warning("excel为空");
                return false;
            }
            // 找出Parameter Key与Parameter Name所在的columIndex
            Sheet sheet = wb.getSheetAt(0);
            Row row0 = sheet.getRow(0);
            int keyColIndex = Integer.MAX_VALUE;
            int valueColIndex = Integer.MAX_VALUE;
            for (Cell c : row0){
                if (c.getStringCellValue().equalsIgnoreCase("Parameter Key")){
                    keyColIndex = c.getColumnIndex();
                }else if (c.getStringCellValue().equalsIgnoreCase("Parameter Name")){
                    valueColIndex = c.getColumnIndex();
                }
            }
            if (keyColIndex == Integer.MAX_VALUE){
                logger.warning("表中不存在Parameter Key");
                return false;
            }
            if (valueColIndex == Integer.MAX_VALUE){
                logger.warning("表中不存在Parameter Name");
                return false;
            }
            // 将对应的Parameter Key与Parameter Name写入properties文件
            for (int i=1;i<=sheet.getLastRowNum();i++){
                Row r = sheet.getRow(i);
                String keyStr = r.getCell(keyColIndex).getStringCellValue();
                String valueStr = r.getCell(valueColIndex).getStringCellValue();
                prop.setProperty(keyStr,valueStr);
            }
            prop.store(propOutputStream,"generate properties form excel");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try{
                if (excelInputStream != null){
                    excelInputStream.close();
                }
                if (propInputStream != null){
                    propInputStream.close();
                }
                if (propOutputStream != null){
                    propOutputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return true;
    }
}
