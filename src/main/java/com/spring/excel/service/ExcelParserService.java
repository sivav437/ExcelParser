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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.spring.excel.model.Customer;
import com.spring.excel.model.ExcelSheet;
import com.spring.excel.model.Product;
import com.spring.excel.repository.CommonRepo;
import com.spring.excel.repository.CustomerRepo;
import com.spring.excel.repository.ProductRepo;

@Service 
public class ExcelParserService {
	
	@Autowired
	private CustomerRepo cust_repo;
	
	@Autowired
	private ProductRepo product_repo;
	
	private CommonRepo repo;
	
	
	
	List<Customer> customers=new ArrayList<>();
    
    List<Product> products=new ArrayList<>();
	
	public Customer cust;
	
	public Product prod;
	
	
	
	public List<List<ExcelSheet>> parseExcelFile(String filePath) throws IOException {
		
		List<List<ExcelSheet>> allData = new ArrayList<>();
		
//        List<List<String>> data = new ArrayList<>();
        
        ClassPathResource resource = new ClassPathResource(filePath);
        
        
        
        
        try (InputStream inputStream = resource.getInputStream();
                Workbook workbook = WorkbookFactory.create(inputStream)) {
        	
        	
               // Get the first worksheet
        	
        	int num_sheets=workbook.getNumberOfSheets();
        	int idx=0;
        	System.out.println(num_sheets+" no. of sheets");
        	int id=1;
        	
        	while(idx<num_sheets) {
        		
        		
        	
               Sheet sheet = (Sheet) workbook.getSheetAt(idx);
               
               String sheet_name=sheet.getSheetName();
               
               System.out.println(sheet_name);
               
               switch(sheet_name) {
               		case "customerData":
               			this.cust=new Customer();
               			System.out.println("customerData sheet switch");
               			this.prod=null;
               			//repo=cust_repo;
               			break;
               		case "productData":
               			System.out.println("productData sheet switch");
               			this.prod=new Product();
               			this.cust=null;
               			//repo=product_repo;

               			break;
               }

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
                       System.out.println("Header type is "+cell.getStringCellValue());
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
                	  
        		  case BLANK:
        			  headers[start_cell_header]= "";
                  
        		  default:
                	  headers[start_cell_header]= "";
        		  }
         	  }
         	  
         	 for(String i:headers) {
         		 System.out.println(i+" Header Name with sheet "+sheet_name);
         	 }
         	  
         	   
               
               for(;start_row<=end_row;start_row++) {
            	   if( !(start_row>0)) { // row 0 consists of headers
            		   continue;
            	   }
            	   
            	   Row row=sheet.getRow(start_row);
            	  
            	   int  start_cell=row.getFirstCellNum();
            	  int  end_cell=row.getLastCellNum();
            	  
            	  Customer localCust = null;
            	  Product localProd = null;
            	  
            	  System.out.println(sheet_name+" sheet name at start_row for loop");
            	  if(sheet_name.equals("customerData")) {
            		  System.out.println("entered at sheet_name with cdata");
            		  localCust=new Customer();
            	  }else if(sheet_name.equals("productData")) {
            		  System.out.println("entered at sheet_name with pdata");
            		  localProd=new Product();
            	  }
            	  System.out.println((localCust==null)+" Local Cust is Null");
            	  for(;start_cell<end_cell;start_cell++) {
            		  
            		  Cell cell=row.getCell(start_cell);
            		  System.out.println(cell.getCellType());
            		  
            		  
            		  
            		  switch(headers[start_cell]) {
            		  case "customer_id":
            			  System.out.println("at customer_id cell of switch "+start_cell);
            			  int customer_id=(int) getCellValue(cell,"customer_id");
            			  localCust.setCustomer_id(id);
            			  id++;
            			  break;
            		  case "customer_name":
            			  System.out.println("at customer_name cell of switch "+start_cell);
            			  String customer_name=(String) getCellValue(cell,"customer_name");
            			  localCust.setCustomer_name(customer_name);
            			  break;
            		  case "customer_password":
            			  System.out.println("at customer_password cell of switch "+start_cell);
            			  String customer_password=(String) getCellValue(cell,"customer_password");
            			  localCust.setCustomer_password(customer_password);
            			  break;
            			  
            		  case "product_id":
            			  int product_id=(int) getCellValue(cell,"product_id");
            			  localProd.setProduct_id(product_id);
            			  break;
            		  case "product_name":
            			  String product_name=(String) getCellValue(cell,"product_name");
            			  localProd.setProduct_name(product_name);
            			  break;
            		  case "product_price":
            			  int product_price=(int) getCellValue(cell,"product_price");
            			  localProd.setProduct_price(product_price);
            			  break;
            		  }
            		  String cellValue = getCellValueAsString(cell); //converting all in a string.
            		  //System.out.print(cellValue+" ");
            	  }
            	  
            	  System.out.println("----------------------------");
            	  if(localCust != null && localCust.getCustomer_id()>0) {
//            		  cust_repo.save(cust);
            		  //Customer cust_copy=cust.clone();
            		  customers.add(localCust);
            		 
            		  System.out.println(localCust+" Customer Data is");
            	  }
            	  
            	  if(localProd !=null && localProd.getProduct_id()>0 ) {
            		  //Product prod_copy=prod.clone();
            		  products.add(localProd);
            		  System.out.println(localProd+" Product Data is");
            	  }
            	  
            	  
            	  
            	   
               }
               
//               sheet.forEach(row -> {
//                   List<String> rowData = new ArrayList<>();

                   // Iterate over cells in the row
                   
//                   Iterator<Cell> cellIterator = ((Row) row).cellIterator();
//                   if(row.getRowNum()==0) {
//                	   System.out.println("GetRowNum is 0");
//                   }
//                   System.out.println("Row nums are "+row.getRowNum()); // starts rows from 0 to value presented.
//                   System.out.println(row.getPhysicalNumberOfCells()+" physical num of cells"); // it gives total cells in a row
//                   System.out.println(row.getLastCellNum()+" laast cell number");
//                   System.out.println(row.getFirstCellNum()+" first cell number");
                  
                   
                   
                   
//                   
//                   while (cellIterator.hasNext()) {
//                	   
//                       Cell cell = cellIterator.next();
//                       
//                       System.out.println(cell.getRowIndex()+"-- cell.getRowIndex");
//                       System.out.println(cell.getColumnIndex()+"-- cell.getcolumnidx");
//                       String cellValue = getCellValueAsString(cell);
                       
                       
//                       if(cell.getColumnIndex()==0) {
//                    	   System.out.println(cell.getCellType()+" cell type ");
//                       }
                       
//                       rowData.add(cellValue);
                       
//                   }
//                   for (Cell cell : row) {
//                       String cellValue = getCellValueAsString(cell);
//                       rowData.add(cellValue);
//                   }
//
//                   
//               });
               System.out.println(idx+"idx is");
               
               idx++;
        } //while block
               
        if(customers!=null) { 
        	
        	allData.add(new ArrayList<ExcelSheet>(customers));
        	System.out.print("customers: "+customers);
        	cust_repo.saveAll(customers);
        	System.out.println(cust_repo.findAll()+" customers from db");
        	}
        
        if(products != null) {
        	allData.add(new ArrayList<ExcelSheet>(products));
        	System.out.print("products: "+products);
        	product_repo.saveAll(products);
        	System.out.println(product_repo.findAll()+" products from db");
        	}
        
        
        //System.out.println("products  : "+products);
           return allData;
        
	} // try block
        
	}  // method block
	
	
	
	private Object getCellValue(Cell cell,String headerName) {
		
		if (cell == null) {
            return null;
        }
		switch(cell.getCellType()) {
		
		 case STRING:
             return cell.getStringCellValue();
		 case NUMERIC:
             if (DateUtil.isCellDateFormatted(cell)) {
                 return cell.getDateCellValue();
             }
             if(headerName=="customer_id" || headerName=="product_id"  || headerName=="product_price") {
            	 return (int)cell.getNumericCellValue();
             }
             return cell.getNumericCellValue();
		 case BOOLEAN:
             return String.valueOf(cell.getBooleanCellValue());
		 case FORMULA:
             return cell.getCellFormula();
		 case BLANK:
			 return null;
         default:
        	 return null;
		
		}

		
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
	
	

