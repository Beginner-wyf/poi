package com.wyf;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.*;

/**
 * @author wangyifan
 * @create 2021/3/23 10:52
 */
public class ExcelWrite {

    public final String PATH = "E:\\ZJIPST_PROJECT\\learning\\excel\\excel-poi\\src\\main\\resources\\";

    @Test
    public void write() throws Exception {
        //1、创建excel 03版工作簿
        Workbook workbook03 = new HSSFWorkbook();
        //Workbook workbook07 = new XSSFWorkbook();
        //2、创建工作表,给表名赋值
        Sheet sheet = workbook03.createSheet("小王的测试表");
        //3、创建第一行
        Row row = sheet.createRow(0);
        //4、创建列
        //第一列
        Cell cellA1 = row.createCell(0);
        cellA1.setCellValue("小王的女朋友");
        //第二列
        Cell cellB1 = row.createCell(1);
        cellB1.setCellValue("小王的媳妇儿");
        //第仨列
        Cell cellC1 = row.createCell(2);
        cellC1.setCellValue("小王的崽崽子");
        //第四列
        Cell cellD1 = row.createCell(3);
        cellD1.setCellValue("日期");

        //3、创建第一行
        Row row2 = sheet.createRow(1);
        //4、创建列
        //第一列
        Cell cellA2 = row2.createCell(0);
        cellA2.setCellValue("雯雯崽崽");
        //第二列
        Cell cellB2 = row2.createCell(1);
        cellB2.setCellValue("小雯子");
        //第仨列
        Cell cellC2 = row2.createCell(2);
        cellC2.setCellValue("雯雯子");
        //第四列
        Cell cellD2 = row2.createCell(3);
        //joda时间工具类
        String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cellD2.setCellValue(time);
        //创建文件流
        FileOutputStream fileOutputStream = null;
        try {
            //生成一张表，IO流操作
            fileOutputStream = new FileOutputStream(PATH + "小王学习日记03.xls");
            //输出
            workbook03.write(fileOutputStream);
            System.out.println("创建完成");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            //关闭流
            assert fileOutputStream != null;
            fileOutputStream.close();
        }
    }

    /**
     * 大数据量写入
     * @throws Exception IO流异常
     */
    @Test
    public void biggerWrite() throws Exception {

        //计算开始时间
//        long start = System.currentTimeMillis();
        //创建工作簿
//        Workbook workbook = new XSSFWorkbook();
        //创建工作表
//        Sheet sheet = workbook.createSheet("小王的大数据");


        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\wangyifan\\Desktop\\指令表.xlsx");
        //获得工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //获得工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        //获取单元格样式对象
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //循环写入行和列
        for (int rowNum = 2; rowNum < 65537; rowNum++) {
            Row row = sheet.createRow(rowNum);
            //设置行高
            row.setHeightInPoints((float) 23.75);
            for (int cellNum = 0; cellNum < 6; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
                cell.setCellStyle(cellStyle);
            }
        }
        System.out.println("大数据插入完成");
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream("C:\\Users\\wangyifan\\Desktop\\指令表.xlsx");
            workbook.write(fileOutputStream);
//            long end = System.currentTimeMillis();
//            System.out.println("完成时间："+(double) (end - start) / 1000);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            assert fileOutputStream != null;
            fileOutputStream.close();
        }
    }

    /**
     * super版大数据量写入
     * @throws Exception IO流异常
     */
    @Test
    public void biggerWriteSuper() throws Exception {
        //计算开始时间
        long start = System.currentTimeMillis();
        //创建工作簿
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        //创建工作表
        Sheet sheet = workbook.createSheet("小王的大数据");
        //循环写入行和列
        for (int rowNum = 0; rowNum < 65537; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("大数据插入完成");
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(PATH + "小王的大数据测试07升级版.xlsx");
            workbook.write(fileOutputStream);
            long end = System.currentTimeMillis();
            System.out.println("完成时间："+(double) (end - start) / 1000);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            assert fileOutputStream != null;
            fileOutputStream.close();
            //清处临时文件
            workbook.dispose();
        }
    }

}
