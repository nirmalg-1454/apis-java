package com.example.newSpring.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;

@Service
public class ExcelInitializationService {

    @Value("${excel.file.path:database.xlsx}")
    private String excelFilePath;

    private static final String[] REQUIRED_HEADINGS = {"ID", "Name", "Email"};

    @PostConstruct
    public void initialize() {
        File file = new File(excelFilePath);
        if (!file.exists() || !hasRequiredHeadings(file)) {
            try {
                initializeExcelFile(file);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private boolean hasRequiredHeadings(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                return false;
            }

            Set<String> headings = new HashSet<>();
            for (Cell cell : headerRow) {
                headings.add(cell.getStringCellValue().trim());
            }

            for (String requiredHeading : REQUIRED_HEADINGS) {
                if (!headings.contains(requiredHeading)) {
                    return false;
                }
            }

            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    private void initializeExcelFile(File file) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < REQUIRED_HEADINGS.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(REQUIRED_HEADINGS[i]);
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } finally {
            workbook.close();
        }
    }
}
