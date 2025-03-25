package com.spring.excel.service;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

//import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelParserService {
	
	public List<List<String>> parseExcelFile(String filePath) throws IOException {
        List<List<String>> data = new ArrayList<>();
        ClassPathResource resource = new ClassPathResource(filePath);
        
        try (InputStream inputStream = resource.getInputStream();
                Workbook workbook = WorkbookFactory.create(inputStream)) {

               // Get the first worksheet
               Sheet sheet = (Sheet) workbook.getSheetAt(0);

               // Iterate over rows
               
               sheet.forEach(row -> {
                   List<String> rowData = new ArrayList<>();

                   // Iterate over cells in the row
                   
                   Iterator<Cell> cellIterator = ((Row) row).cellIterator();
                   
                   while (cellIterator.hasNext()) {
                       Cell cell = cellIterator.next();
                       String cellValue = getCellValueAsString(cell);
                       rowData.add(cellValue);
                   }
//                   for (Cell cell : row) {
//                       String cellValue = getCellValueAsString(cell);
//                       rowData.add(cellValue);
//                   }

                   data.add(rowData);
               });
        }
               


           return data;
        
	}
	
	private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

	public List<List<String>> parseExcelFromUpload(MultipartFile file) throws IOException{
		
		List<List<String>> data = new ArrayList<>();
		
		try (InputStream inputStream = file.getInputStream();
	             Workbook workbook = WorkbookFactory.create(inputStream)) {

	            // Get the first worksheet
	            Sheet sheet = workbook.getSheetAt(0);
	           

	            // Use forEach to iterate over rows
	            sheet.forEach(row -> {
	                List<String> rowData = new ArrayList<>();

	                // Use cellIterator to iterate over cells in the row
	                Iterator<Cell> cellIterator = row.cellIterator();
	                while (cellIterator.hasNext()) {
	                    Cell cell = cellIterator.next();
	                    String cellValue = getCellValueAsString(cell);
	                    rowData.add(cellValue);
	                }

	                data.add(rowData);
	            });
	        }

	        return data;
	    }
	
	
}
