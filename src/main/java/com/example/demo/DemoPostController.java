package com.example.demo;

import org.apache.logging.log4j.util.StringBuilderFormattable;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import io.swagger.annotations.ApiOperation;

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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


@RestController

@RequestMapping(value = "/", produces = MediaType.APPLICATION_JSON_VALUE)
public class DemoPostController {

    @GetMapping(value = "/ping")
    public String getPing(){
        String response_val = "pong";
        System.out.println("Ping has been occurred");
        //Business Logic 추가
        return response_val;
    }

    @GetMapping(value = "/excelDownload")
    public void excelDownload(HttpServletResponse response) throws IOException{

        List<Map<String, Object>> datas = new ArrayList<>();
        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        System.out.println(timestamp+" Excel Download Called by User");

        int i =0;
        int testSize = 100000;

        while(i < testSize){
            Map<String, Object> data1 = new HashMap<>();
            String test_value_a = String.format("%d_test_value_a",i);
            String test_value_b = String.format("%d_test_value_b",i);
            String test_value_c = String.format("%d_test_value_c",i);
            String test_value_d = String.format("%d_test_value_d",i);
            String test_value_e = String.format("%d_test_value_e",i);
            String test_value_f = String.format("%d_test_value_f",i);
            String test_value_g = String.format("%d_test_value_g",i);
            String test_value_h = String.format("%d_test_value_h",i);
            String test_value_i = String.format("%d_test_value_i",i);
            String test_value_j = String.format("%d_test_value_j",i);
            String test_value_k = String.format("%d_test_value_k",i);
            String test_value_l = String.format("%d_test_value_l",i);
            data1.put("a", test_value_a);
            data1.put("b", test_value_b);
            data1.put("c", test_value_c);
            data1.put("d", test_value_d);
            data1.put("e", test_value_e);
            data1.put("f", test_value_f);
            data1.put("g", test_value_g);
            data1.put("h", test_value_h);
            data1.put("i", test_value_i);
            data1.put("j", test_value_j);
            data1.put("k", test_value_k);
            data1.put("l", test_value_l);

            datas.add(data1);
            
            i++;
        }

        ExcelUtil excelUtil = new ExcelUtil();
        excelUtil.createExcelToResponse(
            datas,
            String.format("%s-%s", "data", LocalDate.now().toString()),
            response
        );
    }

} 