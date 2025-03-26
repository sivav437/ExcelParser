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

import com.spring.excel.model.Customer;

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
               int start_row=sheet.getFirstRowNum();
               int end_row=sheet.getLastRowNum();
               
               Row row_header=sheet.getRow(0);
               String[] headers=new String[row_header.getPhysicalNumberOfCells()];
               int  start_cell_header=row_header.getFirstCellNum();
         	   int  end_cell_header=row_header.getLastCellNum();
         	   
         	  for(;start_cell_header<end_cell_header;start_cell_header++) {
        		  Cell cell=row_header.getCell(start_cell_header);
        		  switch(cell.getCellType()) {
        		  case STRING:
                       headers[start_cell_header]= cell.getStringCellValue();
                       System.out.println("Header type is"+"String");
                       break;
                  case NUMERIC:
                      if (DateUtil.isCellDateFormatted(cell)) {
                    	  headers[start_cell_header]= cell.getDateCellValue().toString();
                    	  break;
                      }
                      headers[start_cell_header]= String.valueOf(cell.getNumericCellValue());
                      break;
                  case BOOLEAN:
                	  headers[start_cell_header]= String.valueOf(cell.getBooleanCellValue());
                	  break;
                  case FORMULA:
                	  headers[start_cell_header]= cell.getCellFormula();
                	  break;
                  default:
                	  headers[start_cell_header]= "";
        		  }
         	  }
         	  
         	 for(String i:headers) {
         		 System.out.println(i+" Header Name");
         	 }
         	  
         	   
               
               for(;start_row<=end_row;start_row++) {
            	   if( !(start_row>0)) { // row 0 consists of headers
            		   continue;
            	   }
            	   Row row=sheet.getRow(start_row);
            	  int  start_cell=row.getFirstCellNum();
            	  int  end_cell=row.getLastCellNum();
            	  
            	  for(;start_cell<end_cell;start_cell++) {
            		  Cell cell=row.getCell(start_cell);
            		  System.out.println(cell.getCellType());
            		  
            		  switch(headers[start_cell]) {
            		  case "customer_id":
            			  System.out.println("at customer_id cell of switch "+start_cell);
            			  break;
            		  case "customer_name":
            			  System.out.println("at customer_name cell of switch "+start_cell);
            			  break;
            		  case "customer_password":
            			  System.out.println("at customer_password cell of switch "+start_cell);
            			  break;
            		  }
            		  String cellValue = getCellValueAsString(cell); //converting all in a string.
            		  //System.out.print(cellValue+" ");
            	  }
            	  System.out.println("----------------------------");
            	   
               }
               
               sheet.forEach(row -> {
                   List<String> rowData = new ArrayList<>();

                   // Iterate over cells in the row
                   
                   Iterator<Cell> cellIterator = ((Row) row).cellIterator();
//                   if(row.getRowNum()==0) {
//                	   System.out.println("GetRowNum is 0");
//                   }
//                   System.out.println("Row nums are "+row.getRowNum()); // starts rows from 0 to value presented.
//                   System.out.println(row.getPhysicalNumberOfCells()+" physical num of cells"); // it gives total cells in a row
//                   System.out.println(row.getLastCellNum()+" laast cell number");
//                   System.out.println(row.getFirstCellNum()+" first cell number");
                  
                   
                   
                   
                   
                   while (cellIterator.hasNext()) {
                	   
                       Cell cell = cellIterator.next();
//                       
//                       System.out.println(cell.getRowIndex()+"-- cell.getRowIndex");
//                       System.out.println(cell.getColumnIndex()+"-- cell.getcolumnidx");
                       String cellValue = getCellValueAsString(cell);
                       
                       
//                       if(cell.getColumnIndex()==0) {
//                    	   System.out.println(cell.getCellType()+" cell type ");
//                       }
                       
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
		List<Customer> customers=new ArrayList<>();
		
		try (InputStream inputStream = file.getInputStream();
	             Workbook workbook = WorkbookFactory.create(inputStream)) {

	            // Get the first worksheet
	            Sheet sheet = workbook.getSheetAt(0);
	           

	            // Use forEach to iterate over rows
	            sheet.forEach(row -> {
	                List<String> rowData = new ArrayList<>();

	                // Use cellIterator to iterate over cells in the row
	                Iterator<Cell> cellIterator = row.cellIterator();
	                System.out.println("--> ");
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
