package com.example.importexcel.service;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Service
public class ExcelServiceImpl implements ExcelService {

//    @Override
//    public String excelImport(MultipartFile file) {
//        //excel.XLS文件
//        //HSSFWorkbook hssfWorkbook = null;
//        //excel.XLSX文件
//        XSSFWorkbook xssfWorkbook = null;
//        try {
//            InputStream inputStream = file.getInputStream();
//            xssfWorkbook = new XSSFWorkbook(inputStream);
//        } catch (IOException e) {
////            log.info("创建文件输入流失败：" + e.getMessage());
//            return "创建文件输入流失败";
//        }
//        // 获取Excel的第一个sheet
//        XSSFSheet sheetAt = xssfWorkbook.getSheetAt(0);
//        //获取行数
//        int columnNum = sheetAt.getPhysicalNumberOfRows();
//        for (int i = 0; i < columnNum; i++) {
//            //获取每行
//            Row row = sheetAt.getRow(i);
//            //获取列数
//            int lastRowNum = row.getPhysicalNumberOfCells();
//            for (int j = 0; j < lastRowNum; j++) {
//                //获取每列
//                Cell cell = row.getCell(j);
//                //第i行第j列的值(模板用string数值，如果用其他类型则用其他方法获取值)
//                String cellValue = cell.getStringCellValue();
//                System.out.println("第" + i + "行第" + j + "列数值为:" + cellValue);
//            }
//        }
//        return "success";
//    }

    @Override
    public String excelImport(MultipartFile file) {
        //excel.XLS文件
        HSSFWorkbook hssfWorkbook = null;
        //excel.XLSX文件
//        XSSFWorkbook xssfWorkbook = null;
        try {
            InputStream inputStream = file.getInputStream();
            hssfWorkbook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
//            log.info("创建文件输入流失败：" + e.getMessage());
            return "创建文件输入流失败";
        }
        // 获取Excel的第一个sheet
        HSSFSheet sheetAt = hssfWorkbook.getSheetAt(0);
        //获取行数
        int columnNum = sheetAt.getPhysicalNumberOfRows();
        for (int i = 0; i < columnNum; i++) {
            //获取每行
            Row row = sheetAt.getRow(i);
            //获取列数
            int lastRowNum = row.getPhysicalNumberOfCells();
            for (int j = 0; j < lastRowNum; j++) {
                //获取每列
                Cell cell = row.getCell(j);
                //第i行第j列的值(模板用string数值，如果用其他类型则用其他方法获取值)
                String cellValue = StringUtils.isBlank(cell.getStringCellValue()) ?
                        "" : cell.getStringCellValue();
                System.out.println("第" + i + "行第" + j + "列数值为:" + cellValue);
            }
        }
        return "success";
    }
}
