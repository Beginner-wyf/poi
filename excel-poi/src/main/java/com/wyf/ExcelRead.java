package com.wyf;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

/**
 * @author wangyifan
 * @create 2021/3/23 15:30
 */
public class ExcelRead {

    public final String PATH = "E:\\ZJIPST_PROJECT\\learning\\excel\\excel-poi\\src\\main\\resources\\";

    @Test
    public void read() throws IOException {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(PATH + "小王学习日记07.xlsx");
            //获得工作簿
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            //获得工作表
            XSSFSheet sheet = workbook.getSheet("小王的测试表07");
            //获得行
            XSSFRow row = sheet.getRow(1);
            //获得列
            XSSFCell cell = row.getCell(3);
            System.out.println(cell.getNumericCellValue());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            assert fileInputStream != null;
            fileInputStream.close();
        }
    }

    @Test
    public void read2() throws IOException {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(PATH + "小王的表格读取.xlsx");
            //获得工作簿
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            //获得工作表
            XSSFSheet sheet = workbook.getSheetAt(0);
            //获得表格标题行内容
            //getTitle(sheet);

            //读取表中的内容
            //获取有效总行数
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int rowNum = 0; rowNum < rowCount; rowNum++) {
                XSSFRow row = sheet.getRow(rowNum);
                if (row != null){
                    //读取列
                    //获取当前行的总列数（有数据的列数）
                    int cellCount = row.getPhysicalNumberOfCells();
                    for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                        XSSFCell cell = row.getCell(cellNum);
                        if (cell != null){
                            CellType cellType = cell.getCellType();
                            String cellValue = getCellValue(cell, cellType);
                            System.out.print(cellValue + "(" +cellType+")|");
                        }
                    }
                }
                System.out.println();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            assert fileInputStream != null;
            fileInputStream.close();
        }
    }

    /**
     * 读取公式
     */
    @Test
    public void getFormula() throws Exception{
        //读取文件
        FileInputStream fileInputStream = new FileInputStream(PATH + "公式计算.xlsx");
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表
        Sheet sheet = workbook.getSheetAt(0);

        //获取行列
        Row row = sheet.getRow(sheet.getPhysicalNumberOfRows() - 1);
        Cell cell = row.getCell(row.getPhysicalNumberOfCells()-1);

        //拿到计算公式
        FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator(workbook);

        //输出公式
        CellType cellType = cell.getCellType();
        String cellValue = getCellValue(cell, cellType);
        System.out.println(cellValue);

        //获取该公式计算的值
        CellValue evaluate = formulaEvaluator.evaluate(cell);
        String res = evaluate.formatAsString();
        System.out.println(res);
    }

    private String getCellValue(Cell cell, CellType cellType) {
        String cellValue;
        switch (cellType){
            //没有值
            case _NONE:
                cellValue = "none";
                break;
            case NUMERIC:
                //若果是日期格式
                if (HSSFDateUtil.isCellDateFormatted(cell)){
                    cellValue = new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd HH:mm:ss");
                }else {
                    //数字格式,防止数字过长，可以将其转化成字符串再侠士
                    //cell.setCellType(CellType.STRING);cell.toString;
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case FORMULA:
                cellValue = cell.getCellFormula();
                break;
            case BLANK:
                cellValue = "空";
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case ERROR:
                cellValue = "Error";
                System.out.println("表格读取错误");
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    private void getTitle(XSSFSheet sheet) {
        XSSFRow row = sheet.getRow(0);
        if (row != null){
            //获取当前行的总列数（有数据的列数）
            int cellCount = row.getPhysicalNumberOfCells();
            ArrayList<String> strings = new ArrayList<>();
            for (int i = 0; i < cellCount; i++) {
                XSSFCell cell = row.getCell(i);
                if (cell != null){
                    CellType cellType = cell.getCellType();
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + "(" +cellType+")|");
                }
            }
        }
    }


}
