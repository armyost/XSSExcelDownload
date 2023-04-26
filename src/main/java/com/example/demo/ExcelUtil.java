package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import java.util.List;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.sql.Timestamp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
    
    private int rowNum = 0;
    
    //File로 만들 경우
    public void createExcelToFile(List<Map<String, Object>> datas, String filepath) throws FileNotFoundException, IOException {
        //workbook = new HSSFWorkbook(); // 엑셀 97 ~ 2003
        //workbook = new XSSFWorkbook(); // 엑셀 2007 버전 이상

        XSSFWorkbook workbook = new XSSFWorkbook(); 
        XSSFSheet sheet = workbook.createSheet("XSSFWorkBook Normal");

        rowNum = 0;

        createExcel(sheet, datas);


        FileOutputStream fos = new FileOutputStream(new File(filepath));
        workbook.write(fos);
        workbook.close();

    }
    
    //HttpServletResponse 경우
    public void createExcelToResponse(List<Map<String, Object>> datas, String filename, HttpServletResponse response) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(); // 성능 개선 버전
        XSSFSheet sheet = workbook.createSheet("데이터");
        

        rowNum = 0;
        Timestamp timestamp_a = new Timestamp(System.currentTimeMillis());
        System.out.println(timestamp_a+" Excel export by HTTP Started");

        createExcel(sheet, datas);
        
        // 컨텐츠 타입과 파일명 지정
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", String.format("attachment;filename=%s.xlsx", filename));
        
        workbook.write(response.getOutputStream());
        workbook.close();
        Timestamp timestamp_b = new Timestamp(System.currentTimeMillis());
        System.out.println(timestamp_b+" Excel export by HTTP Finished");
    }

    //엑셀 생성
    private void createExcel(XSSFSheet sheet, List<Map<String, Object>> datas) {
        
        Timestamp timestamp_c = new Timestamp(System.currentTimeMillis());
        int memCheckInterval = 1000;
        int i=0;

        System.out.println(timestamp_c+" Excel create Started");
    
        //데이터를 한개씩 조회해서 한개의 행으로 만든다.
        for (Map<String, Object> data : datas) {
            //row 생성
            Row row = sheet.createRow(rowNum++);
            int cellNum = 0;
            
            //map에 있는 데이터를 한개씩 조회해서 열을 생성한다.
            for (String key : data.keySet()) {
                //cell 생성
                Cell cell = row.createCell(cellNum++);
               	
                //cell에 데이터 삽입
                cell.setCellValue(data.get(key).toString());

            }
            if(i%memCheckInterval == 0) {
                long heapSize = Runtime.getRuntime().totalMemory();
                Timestamp timestamp_e = new Timestamp(System.currentTimeMillis());
                System.out.println("["+i+"] heapSize is "+heapSize+" "+timestamp_e);
            }
            i++;
        }
        Timestamp timestamp_d = new Timestamp(System.currentTimeMillis());
        System.out.println(timestamp_d+" Excel create Finished");
    }
}
