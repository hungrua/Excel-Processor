package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws IOException {
        FileProcesserService processerService = new FileProcesserService();
//        String srcPath = "D:\\ExcelFaceNet\\Master_data_APS.xlsx";
//        String tartPath = "D:\\ExcelFaceNet\\Master_data_APS_new.xlsx";
//        processerService.processing(srcPath, tartPath);
        JFrame frame = new JFrame("My Swing App");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 300);

        JLabel label = new JLabel("Xin chào Java Swing!", SwingConstants.CENTER);
        frame.add(label);

        frame.setVisible(true);
    }
}