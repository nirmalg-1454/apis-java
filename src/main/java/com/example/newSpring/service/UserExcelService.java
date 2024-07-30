package com.example.newSpring.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class UserExcelService {

    @Value("${excel.file.path}")
    private String excelFilePath;

    public String readExcelToJson() throws IOException {
        List<Map<String, String>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row currentRow = sheet.getRow(rowIndex);
                if (currentRow == null) continue;

                Map<String, String> rowData = new HashMap<>();
                for (int colIndex = 0; colIndex < headerRow.getLastCellNum(); colIndex++) {
                    Cell headerCell = headerRow.getCell(colIndex);
                    Cell currentCell = currentRow.getCell(colIndex);

                    String header = headerCell.getStringCellValue().trim();
                    String cellValue = "";

                    if (currentCell != null) {
                        switch (currentCell.getCellType()) {
                            case STRING:
                                cellValue = currentCell.getStringCellValue();
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(currentCell)) {
                                    cellValue = currentCell.getDateCellValue().toString();
                                } else {
                                    // Convert numeric value to string
                                    cellValue = Double.toString(currentCell.getNumericCellValue());
                                }
                                break;
                            case BOOLEAN:
                                cellValue = Boolean.toString(currentCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                // Handle formula cells
                                cellValue = currentCell.getCellFormula();
                                break;
                            case BLANK:
                                cellValue = "";
                                break;
                            default:
                                cellValue = "";
                        }
                    }

                    rowData.put(header, cellValue);
                }
                data.add(rowData);
            }
        }

        ObjectMapper objectMapper = new ObjectMapper();
        return objectMapper.writeValueAsString(data);
    }

    public void writeDataToExcel(String excelFilePath, List<Map<String, Object>> data) throws IOException {
        try (FileInputStream fis = new FileInputStream(this.excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Determine the starting row for appending new data
            int lastRowNum = sheet.getLastRowNum();

            // Append data starting from the next row after the last existing row
            int startRowIndex = lastRowNum + 1;

            // Assume data is non-empty and headers are present
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Header row is missing.");
            }

            // Write new data
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(startRowIndex + i);
                Map<String, Object> rowData = data.get(i);

                int colIndex = 0;
                for (Cell headerCell : headerRow) {
                    String header = headerCell.getStringCellValue().trim();
                    Cell cell = row.createCell(colIndex++);
                    Object value = rowData.get(header);

                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    } else {
                        cell.setCellValue(value != null ? value.toString() : "");
                    }
                }
            }

            // Write changes to file
            try (FileOutputStream fos = new FileOutputStream(this.excelFilePath)) {
                workbook.write(fos);
            }
        }
    }

}
