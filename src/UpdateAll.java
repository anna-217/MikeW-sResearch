import java.io.*;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
/*
 * 				 cell=2 date column
 *				 cell = 1 id
 *				 cell = 0 hyperlink
 */

public class UpdateAll {
	
	public static void main(String[] args) {
		UpdateAll up = new UpdateAll();
		
		Scanner scann = new Scanner(System.in);
		System.out.println("\nPlease back up the file first.\n\nEnter file name (without file extension):");
			String fname = scann.next();
			System.out.println("\nEnter master data sheet name:");
			String mname = scann.next();
			System.out.println("\nEnter new data sheet name:"	);
			String nname = scann.next();
			scann.close();
			fname = fname + ".xlsx";
			
			System.out.println("\n*************************\n");
			
			up.updateDate(fname, mname, nname);
	} // end of main function
		

	public void updateDate (String filename, String sheetName, String newSheetName) {
		Map <String, CellAddress> masterMap = new HashMap<String, CellAddress>();
		Map <String, CellAddress> newMap = new HashMap<String, CellAddress>();
		int lastrowNum = 0;
		DataFormatter formatter = new DataFormatter();
		
		try {
//			File file = new File(filename);
//			Workbook workbook = WorkbookFactory.create(file);
			FileInputStream inputFile = new FileInputStream(new File(filename));
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputFile);
			XSSFSheet master = workbook.getSheet(sheetName);
			XSSFSheet newsheet = workbook.getSheet(newSheetName);
			
			// store all the id in master into a map;
			Row row = null;
			Iterator<Row> it = master.iterator();
			if (it.hasNext())
				it.next(); // skip first row
			
			// read id and store cell address in master
			while (it.hasNext()) {
				row = (Row)it.next();
				Cell cell = row.getCell(1);
				if (cell != null){
					String cellValue = formatId(formatter.formatCellValue(cell));
					if (cellValue.length()> 0) {
						if (masterMap.containsKey(cellValue)) {
							System.out.println("!!!!\nFound duplicated id in master. id = "+ cellValue + " at " 
						+ masterMap.get(cellValue)+ " will be ignored.");
						}
						masterMap.put(cellValue, cell.getAddress());
						System.out.println("Reading master cell id:" + cellValue + "**cell address= " + cell.getAddress());
					}
				}
			} 
			
			it = newsheet.iterator();
			if (it.hasNext())
				it.next();
			
			// read id and store cell address in new
			while (it.hasNext()) {
				row = (Row)it.next();
				Cell cell = row.getCell(1);
				if (cell != null){
					String cellValue = formatId(formatter.formatCellValue(cell));
					if(cellValue.length() > 0)
						newMap.put(cellValue, cell.getAddress());
					//if(cellValue.equals("12264???"))
					System.out.println("Reading new-sheet cell id:" + cellValue + "**cell address= " + cell.getAddress());
					
				}
			}
			
			lastrowNum = master.getLastRowNum();
				
		int count = 0;
		// compare element in map
	//	SimpleDateFormat DtFormat = new SimpleDateFormat("dd/MM/yyyy");
		for (String id : masterMap.keySet()) {
			System.out.println("comparing: ID = " + id);
			if (newMap.containsKey(id)) {
				System.out.println("new contains "+ id);
				
				CellAddress madd = masterMap.get(id);
				CellAddress nadd = newMap.get(id);
				
			//	System.out.println("master add = "+madd + " , new add = " + nadd);
				
				Row mrow = master.getRow(madd.getRow());
				
			//	System.out.println("master row number = " + mrow);
				
				Cell mcell = mrow.getCell(2);
				
			//	System.out.println("master date value = " + mcell);
				
				Row nrow = newsheet.getRow(nadd.getRow());
				
			//	System.out.println("new row number = " + nrow);
				Cell ncell = nrow.getCell(2);
			//	System.out.println("new date value = " + ncell);
				
				Cell masterLastCell = mrow.getCell(19);
                Cell newAddCell = nrow.getCell(3);
                
                String master_date = formatter.formatCellValue(mcell);
            
                String new_date = formatter.formatCellValue(ncell);
             //   System.out.println("dates master = "+ master_date + " new = " + new_date);
                String new_address = formatter.formatCellValue(newAddCell);
                

                 if (!master_date.equals(new_date)) {
                     System.out.println("#ID = " + id + " is updated.");
                     count++;
                     mcell.setCellValue(new_date);
                     if (masterLastCell == null)
                         mcell = mrow.createCell(19);
                     else
                         mcell = masterLastCell;
                     
                     mcell.setCellValue(new_address);
                 
                 }
				
				// copy hyperlink
				mcell = mrow.getCell(0);
				ncell = nrow.getCell(0);
				if (ncell.getHyperlink() != null) {
					count++;
					mcell.setHyperlink(ncell.getHyperlink());
					mcell.setCellStyle(ncell.getCellStyle());
				}
				newMap.remove(id);
			}// end if
		}
			
		System.out.println(count + " rows are affected. Finish updating the dates.\n\n******************************");
	
		FileOutputStream output = new FileOutputStream(new File(filename));
		workbook.write(output);
		output.flush();
		output.close();
		workbook.close();
		System.out.println("Found " + newMap.size() + " new records.");
		if (!newMap.isEmpty()) {
			System.out.println(newMap.size() + " new records will be added to master sheet.");
			// row 405
			addRecords(newMap, lastrowNum + 1, filename, sheetName, newSheetName);
		}
		
		
		}catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		System.out.println("DONE!");
	}
	
	@SuppressWarnings({ "deprecation", "deprecation", "deprecation", "deprecation" })
	private void addRecords (Map<String, CellAddress> map, int insertAt, String filename, String masterSheetName, String newSheetName) {
		try {
			
//			File file = new File(filename);
//			Workbook workbook = WorkbookFactory.create(file);
			FileInputStream input = new FileInputStream(new File(filename));
			XSSFWorkbook workbook = new XSSFWorkbook(input);
			XSSFSheet master = workbook.getSheet(masterSheetName);
			XSSFSheet newsheet = workbook.getSheet(newSheetName);			
			DataFormatter formatter = new DataFormatter();
			
			for (String key : map.keySet()) {
				CellAddress add = map.get(key);
				Row newrow = newsheet.getRow(add.getRow());
				Row masterrow = master.createRow(insertAt++);
				for (int i = 0; i < newrow.getLastCellNum(); i++) {
					// iterate through the row, copy style, copy value
					Cell newcell = newrow.getCell(i);
					Cell mascell = masterrow.createCell(i);
					
					 // If the cell in new is null jump to next cell
			        if (newcell == null) {
			            mascell = null;
			            continue;
			        }
					CellStyle style = workbook.createCellStyle();
					style.cloneStyleFrom(newcell.getCellStyle());
					mascell.setCellStyle(style);
					
					// copy cell comment
					if (newcell.getCellComment() != null) {
						mascell.setCellComment(newcell.getCellComment());
					}
					
					// copy hyper-link
					if (newcell.getHyperlink() != null) {
						mascell.setHyperlink(newcell.getHyperlink());
					}
					
					// copy data type
					mascell.setCellType(newcell.getCellTypeEnum());
					
					 // Set the cell data value
			        switch (newcell.getCellType()) {
			            case Cell.CELL_TYPE_BLANK:
			            	mascell.setCellValue(newcell.getStringCellValue());
			                break;
			            case Cell.CELL_TYPE_BOOLEAN:
			            	mascell.setCellValue(newcell.getBooleanCellValue());
			                break;
			            case Cell.CELL_TYPE_ERROR:
			            	mascell.setCellErrorValue(newcell.getErrorCellValue());
			                break;
			            case Cell.CELL_TYPE_FORMULA:
			            	mascell.setCellFormula(newcell.getCellFormula());
			                break;
			            case Cell.CELL_TYPE_NUMERIC:
			            	mascell.setCellValue(newcell.getNumericCellValue());
			                break;
			            case Cell.CELL_TYPE_STRING:
			            	mascell.setCellValue(newcell.getRichStringCellValue());
			                break;
			        }
								
					if (i == 1) {
						System.out.println("Adding data with id = " + newcell);
					}
				}
			}// end of for
			
			FileOutputStream out = new FileOutputStream(new File(filename));
			workbook.write(out);
			out.flush();
			out.close();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public String formatId(String rawId) {
		String formatedId = rawId.replaceAll("\\D+","");
		return formatedId;
	}
}
