package com.spring.excel.controller;

import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.spring.excel.model.Customer;
import com.spring.excel.model.ExcelSheet;
import com.spring.excel.service.ExcelParserService;

@RestController
public class ExcelController {

	private final ExcelParserService excelParserService;

    @Autowired
    public ExcelController(ExcelParserService excelParserService) {
        this.excelParserService = excelParserService;
    }
    
    @GetMapping("/parse-excel")
    public List<List<ExcelSheet>> parseExcel(@RequestParam("classPath") String classPath) throws IOException { //@RequestParam("filePath") String filePath
    	
        return (List<List<ExcelSheet>>) excelParserService.parseExcelFile(classPath);
    }
    
    
    @PostMapping("/upload-excel")
    public ResponseEntity<List<List<ExcelSheet>>> uploadExcel(@RequestParam("file") MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().body(null); 
        }

        List<List<ExcelSheet>> parsedData = excelParserService.parseExcelFromUpload(file);
        return ResponseEntity.ok(parsedData); 
    }
    
}
