package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws IOException {
        FileProcesserService processerService = new FileProcesserService();
        String srcPath = "D:\\ExcelFaceNet\\Master_data_APS.xlsx";
        String tartPath = "D:\\ExcelFaceNet\\Master_data_APS_new.xlsx";
        processerService.processing(srcPath, tartPath);
    }
}