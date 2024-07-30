package com.example.newSpring.controller;

import com.example.newSpring.model.User;
import com.example.newSpring.service.UserExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    @Value("${excel.file.path}")
    private String excelFilePath;
    @Autowired
    private UserExcelService excelService;

    @Autowired
    public ExcelController(UserExcelService excelService){
        this.excelService = excelService;
    }

    @GetMapping("/read")
    public String readExcel() {
        try {
            return excelService.readExcelToJson();
        } catch (IOException e) {
            return "Error reading data: " + e.getMessage();
        }
    }

    @PostMapping("/write")
    public ResponseEntity<String> writeExcel(@RequestBody List<Map<String, Object>> data) {
        try {
            excelService.writeDataToExcel(excelFilePath, data);
            return ResponseEntity.ok("Data written to Excel file successfully.");
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to write data to Excel file: " + e.getMessage());
        }
    }
}
