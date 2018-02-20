package com.dxc.sabrix;


import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.hp.sabrix.SabrixConnector;

public class SabrixClientData {

	        public static void main(String[] args) throws IOException, XMLStreamException {
	        
	            Integer ColumnIndexNumber = null;
	            Integer totalTaxAmountIndex = null;	
	            Integer effTaxRateIndex = null;
	            Integer taxCodeIndex = null;
	            Integer authorityMessagesIndex = null;
	            String onesourceResponse = null;
	            Integer responseXMLIndex = null;
	            Boolean flag=true;  
	            String TOTAL_TAX_AMOUNT= null;
	            String EFFECTIVE_TAX_RATE = null;
	            String AUTHORITY_NAME = null;
	            String RULE_ORDER = null;
	            String JURISDICTION_TEXT= null;
	            String ERP_TAX_CODE = null;
	            String Message= "";

				try {
						FileInputStream fileInputStream = new FileInputStream("C:/Users/rameshb/Desktop/HPE_WorkBenchScenario_2.xlsm");
						XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream); 
						XSSFSheet worksheet = workbook.getSheetAt(0);
						ArrayList<String> columndata = new ArrayList<>();;
						Iterator<Row> rowIterator = worksheet.iterator(); 
						//added for response update in xls
						Iterator<Row> rowIteratorForAddResponse = worksheet.iterator();
				        XSSFRow rowNum = worksheet.getRow(1);
						        while (rowIterator.hasNext()) {
						            Row row = rowIterator.next();
						            Iterator<Cell> cellIterator = row.cellIterator();
						            while (cellIterator.hasNext()) {
						                Cell cell = cellIterator.next();
						                if(cell.getCellType()==Cell.CELL_TYPE_STRING){ 
						                    String  columnName = cell.getStringCellValue();
						                    if ("XML".equalsIgnoreCase(columnName)) {
						                    	 ColumnIndexNumber=cell.getColumnIndex();
						                    	 break;
						                     }                
						                   }
						                if(ColumnIndexNumber!=null) {
						                    if(row.getRowNum() > 0){ 
						                        if(cell.getColumnIndex() == ColumnIndexNumber){
						                            switch (cell.getCellType()) {
						                            case Cell.CELL_TYPE_FORMULA:  //TODO need to change String to formula
						                                columndata.add(cell.getStringCellValue());
						                                break;
						                            }
						                        }
						                    }
						                }
						                
						            }
						        }			    
				            if(columndata.size()>0) {
				            	int condition=columndata.size();
						        for(int j=0;j<condition;j++) {
						            onesourceResponse=generateResponse(columndata.get(j));
						            if(onesourceResponse!=null) {
									byte[] byteArray = onesourceResponse.getBytes("UTF-8");
								    ByteArrayInputStream inputStream = new ByteArrayInputStream(byteArray);
								    XMLInputFactory inputFactory = XMLInputFactory.newInstance();
								    XMLEventReader xmlEventReader = inputFactory.createXMLEventReader(inputStream);
			 			            while(xmlEventReader.hasNext()){
						                XMLEvent xmlEvent = xmlEventReader.nextEvent();
						                if (xmlEvent.isStartElement()){
						                   StartElement startElement = xmlEvent.asStartElement();
						                   if(startElement.getName().getLocalPart() .equals("ERP_TAX_CODE")){
						                	   xmlEvent = xmlEventReader.nextEvent();
						                	   ERP_TAX_CODE=xmlEvent.toString();  
						                   } else if(startElement.getName().getLocalPart().equals("TOTAL_TAX_AMOUNT")){
						                	   xmlEvent = xmlEventReader.nextEvent(); 
						                	  
						                	 if(flag==true) {  
						                		 TOTAL_TAX_AMOUNT=xmlEvent.toString();
						                	 }
						                	  
						                	   flag=false;
						                   } else if(startElement.getName().getLocalPart().equals("EFFECTIVE_TAX_RATE")){
						                	   xmlEvent = xmlEventReader.nextEvent();
						                	   EFFECTIVE_TAX_RATE=xmlEvent.toString()+";"+" ";
						                   }
						                   else if(startElement.getName().getLocalPart().equals("AUTHORITY_NAME")){
						                	   xmlEvent = xmlEventReader.nextEvent();
						                	   AUTHORITY_NAME="AUTHORITY_NAME:"+xmlEvent.toString()+";"+" ";
						                	   Message=Message+AUTHORITY_NAME;
						                   }
						                   else if(startElement.getName().getLocalPart().equals("RULE_ORDER")){
						                	   xmlEvent = xmlEventReader.nextEvent();
						                	   RULE_ORDER="RULE_ORDER:"+xmlEvent.toString()+'\n';
						                	   Message=Message+RULE_ORDER;
						                   }
						                   else if(startElement.getName().getLocalPart().equals("JURISDICTION_TEXT")){
						                	   xmlEvent = xmlEventReader.nextEvent();
						                	   JURISDICTION_TEXT="JURISDICTION_TEXT:"+xmlEvent.toString()+";"+" ";
						                	   Message=Message+JURISDICTION_TEXT;
						                   }
			
						                }
						           		                
						            }
						            // TODO   for update xls with response
						         while (rowIteratorForAddResponse.hasNext()) {
						                Row row = rowIteratorForAddResponse.next();
						                Iterator<Cell> cellIteratorForesponse = row.cellIterator();
						                while (cellIteratorForesponse.hasNext()) {
						                    Cell cell = cellIteratorForesponse.next();
						                    if(cell.getCellType()==Cell.CELL_TYPE_STRING){ 
				                                String  columnName = cell.getStringCellValue();
				                                if ("Total Tax Amount".equalsIgnoreCase(columnName)) {
				                                	totalTaxAmountIndex=cell.getColumnIndex();
				                                	System.out.println(totalTaxAmountIndex);
				                                 }
				                                if ("Eff. Tax Rate".equalsIgnoreCase(columnName)) {
				                                	effTaxRateIndex=cell.getColumnIndex();
					                                 }
				                                if ("Tax Code".equalsIgnoreCase(columnName)) {
				                                	taxCodeIndex=cell.getColumnIndex();
					                                 }
				                                if ("Authority Messages".equalsIgnoreCase(columnName)) {
				                                	authorityMessagesIndex=cell.getColumnIndex();
					                                 }
				                                if ("ResponseXML".equalsIgnoreCase(columnName)) {
				                                	responseXMLIndex=cell.getColumnIndex();
					                                 }
				                               }
						                }
						            }
						                    if(totalTaxAmountIndex!=null) {
						                    	 Cell lastCellInRow =worksheet.getRow(j+3).getCell(worksheet.getRow(j+3).getLastCellNum() - 1);
						                    	 System.out.println(lastCellInRow);
						                    	Cell totalTaxAmountCell = worksheet.getRow(j+3).createCell(totalTaxAmountIndex);
						                            	totalTaxAmountCell.setCellValue(TOTAL_TAX_AMOUNT);
						                    }
						                  	                   
						                    if(effTaxRateIndex!=null) {
						                    	Cell effTaxRateCell = worksheet.getRow(j+3).createCell(effTaxRateIndex);
				                            	effTaxRateCell.setCellValue(EFFECTIVE_TAX_RATE);
						                    	}
						                    if(taxCodeIndex!=null) {
						                    	Cell taxCodeCell = worksheet.getRow(j+3).createCell(taxCodeIndex);
						                        
						                            	taxCodeCell.setCellValue(ERP_TAX_CODE);
						
						                    	}
						                    if(authorityMessagesIndex!=null) {
						                    	Cell authorityMessagesCell = worksheet.getRow(j+3).createCell(authorityMessagesIndex);
						                     
						                        	authorityMessagesCell.setCellValue(Message);
							                    	}
						                    if(responseXMLIndex!=null) {
						                    	Cell responseXMLCell = worksheet.getRow(j+3).createCell(responseXMLIndex);
						                        	responseXMLCell.setCellValue(onesourceResponse);
						                    }
						                    
						        		}    
						            }
				            	 
				            }

			            // downloading xls
 			         FileOutputStream outputStream = new FileOutputStream("C:/Users/rameshb/Desktop/HPE_WorkBenchScenario_2Response.xls");
 			        	  workbook.write(outputStream);	        	 		            
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
  
	        }

			private static String generateResponse(String columndata) {
                String newUrl = "http://gte-dev.itcs.hpecorp.net/sabrix/xmlinvoice"; 
                String sbxPayload = columndata;
               	String response = null;
                if (newUrl != null) {
                    URL hostURL = null;
                    try {
                            hostURL = new URL(newUrl);
                    } catch (MalformedURLException e) {
                            e.printStackTrace();
                    }
                            if (newUrl != null) {
                                    response = SabrixConnector.post(hostURL, sbxPayload);	
                                    
                            }

                }
				return response;
			}

		}