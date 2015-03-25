package com.report.template;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 1. Reading the Report Template
 * 2. Fill in the data according to the variable in the Template
 * 3. Rebuild the Report Frame.
 * 
 * @author Shawn Xue
 *
 */
public class PipelineReportExport extends AbstractReportExport {

	private static final long serialVersionUID = 1L;
	private Workbook wb;
	private Sheet sheet;

	/**
	 * Import the Report Template
	 */
	@Override
	public Workbook readExportTemplate(String inportFileName) {
		wb = null;
		sheet = null;
		FileInputStream fi = null;
		try {
			fi =  new FileInputStream(getReportInfo().getInportFileName());
			wb = WorkbookFactory.create(fi);
  
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally{
			if(fi != null){
				try {
					fi.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}				
		}
		return wb;
	}

	/**
	 * Fill the Report Data
	 */
	@Override
	public void fillInReportData(Workbook wb) {
		for(String key:getReportInfo().getReportData().keySet()){
			sheet = wb.getSheet(key);
			Map<String,Object> detail = getReportInfo().getReportData().get(key);
			fillSheetData(detail);
			
		}

	}
	
	/**
	 * Rebuild the Report Frame
	 */
	@Override
	public void rebuildReportFrame(Workbook wb) {
		 int rowNum = sheet.getLastRowNum();
		 Row row = sheet.createRow(rowNum+1);
		 int firstNum = sheet.getRow(rowNum).getFirstCellNum();
		 int lastNum = sheet.getRow(rowNum).getLastCellNum()-1;
		 for(int i= firstNum;i<= lastNum;i++){
			 CellStyle style = wb.createCellStyle();
			 if(i==firstNum){
				 style.setBorderLeft(CellStyle.BORDER_MEDIUM);				 
			 }
			 else if(i==lastNum){
				 style.setBorderRight(CellStyle.BORDER_MEDIUM);
			 }
			 style.setBorderBottom(CellStyle.BORDER_MEDIUM);
			 row.createCell(i).setCellStyle(style);
		 }
		  
	}

	/**
	 * Fill the Sheet according to the Variable filled
	 * @param detail
	 */
	private void fillSheetData(Map<String,Object> detail) {
		int rowNum = sheet.getPhysicalNumberOfRows();
outer:		for(int i=0;i<rowNum;i++){
			Row row = sheet.getRow(i);
			Iterator<Cell> itCell = row.cellIterator(); 
			while(itCell.hasNext()){
				Cell cell = itCell.next();
				String cellValue = cell.getRichStringCellValue().getString().trim();
				if(cellValue.indexOf("{")>=0){
					if(mapCellValueWithInfo(cellValue,detail,cell)){
						break outer;
					}
				}
				else{
					continue;
				}				
			}
		}		
	}
 

	/**
	 * Fill the Report Data according to the data of the Info Object
	 * @param cellValue
	 * @param detail
	 * @param cell
	 * @return
	 */
	private boolean mapCellValueWithInfo(String cellValue,
			Map<String, Object> detail, Cell cell) {
		if(cellValue.indexOf("List.")<0){
			mapColumnWithInfo(cellValue,detail,cell);// fill the column
			return false;
		}
		else{
			mapIteratorWithInfo(cellValue,detail,cell);// fill the list
			return true;
		}
		
	}

	/**
	 * Fill the Column Value with the property value mapped in the Info
	 * @param cellValue
	 * @param detail
	 * @param cell
	 */
	private void mapColumnWithInfo(String cellValue, Map<String, Object> detail, Cell cell) {
		String value = cellValue.substring(1,cellValue.length()-1);
		if(value.indexOf(":")>0){
			String[] tmpValue = value.split(":"); 
			setCellValue(cell,getDetailValue(detail,tmpValue[0],tmpValue[1]));
			setCellStyle(tmpValue[1],cell);				
		}
		else{ 
			setCellValue(cell,getDetailValue(detail,value,null));
		}		
	}


	/**
	 * Fill in the list with the result from DB
	 * @param cellValue
	 * @param detail
	 * @param cell
	 */
	@SuppressWarnings("unchecked")
	private void mapIteratorWithInfo(String cellValue,
			Map<String, Object> detail, Cell cell) {		 
		List<Map<String,Object>> result = (List<Map<String,Object>>)detail.get("List");
		if(result==null||result.size()==0){
			//if there is no result from DB
			sheet.removeRow(cell.getRow());
		}
		else{
			Row row = cell.getRow();
			int num = row.getRowNum();
			int endNum = num+result.size()+1;
			row = copyRow(row,endNum);
			for(int i=num;i<result.size()+num;i++){
				Row newRow = sheet.createRow(i);
				Map<String,Object> listInfo = result.get(i-num); 
				setCellData(row, newRow, listInfo);			 
			}
			sheet.removeRow(row);
		}		
	}

	/**
	 * Fill the Cell with the result from DB, mapping the column name with the key in mapping
	 * if there is formatter string given, the cell style will be initialized
	 * @param row
	 * @param newRow
	 * @param listInfo
	 */
	private void setCellData(Row row, Row newRow, Map<String, Object> listInfo) {
		Iterator<Cell> itCell = row.cellIterator();
		while(itCell.hasNext()){
			Cell formatCell = itCell.next(); 
			Cell newCell = newRow.createCell(formatCell.getColumnIndex());
			newCell.setCellStyle(formatCell.getCellStyle());					
			String valueStr = formatCell.getRichStringCellValue().getString();
			valueStr = valueStr.substring(1,valueStr.length()-1);
			if(valueStr.indexOf(":")<0){ 
				String keyValue = valueStr.split("\\.")[1];
				setCellValue(newCell,getDetailValue(listInfo,keyValue,null));   
				keyValue = null;						
			}
			else{
				String[] formatKeyValue = valueStr.split(":");
				String formatStr = formatKeyValue[1];
				String keyStr = formatKeyValue[0].split("\\.")[1];
				setCellValue(newCell,getDetailValue(listInfo,keyStr,formatStr));  
				setCellStyle(formatStr,newCell); 						 
				keyStr = null;
				formatStr = null;
			}					
		}
	}
	
	/**
	 * get the value from Map, 
	 * if it is not exited in the Map return null or return the String of the value
	 * @param detail
	 * @param key
	 * @return
	 */
	private Object getDetailValue(Map<String, Object> detail, String key,String format) {
		if(detail.get(key)!=null){
			if(format == null){
				return detail.get(key).toString();
			}
			else if(format.indexOf("0")>=0){
				try{
					return Double.valueOf(detail.get(key).toString());
				}
				catch(Exception e){
					return new Double(0);
				}
			}
			else {
				try{
					return Date.valueOf(detail.get(key).toString());
				}
				catch(Exception e){
					return "";
				}
			}
						
		}
		return "";
	}

	/**
	 * Set Cell Value with Formatter Data
	 * @param newCell
	 * @param value
	 */
	private void setCellValue(Cell newCell, Object value) {
		if(value instanceof Date){
			newCell.setCellValue((Date)value);
		}
		else if(value instanceof Double){
			newCell.setCellValue((Double)value);
		}
		else if(value instanceof String){
			newCell.setCellValue((String)value);
		}		
	}
	

	
	/**
	 * Set the style of the cell according to the parameter
	 * 
	 * @param string
	 * @param cell 
	 */
	private void setCellStyle(String string, Cell cell) {
		CellStyle style = cell.getCellStyle();
		DataFormat formatter = wb.createDataFormat();		
		style.setDataFormat(formatter.getFormat(string));
		cell.setCellStyle(style);
		
	}

	/**
	 * Copy the row to the new row number given,return the new row
	 * @param row
	 * @param endNum
	 * @return
	 */
	private Row copyRow(Row row, int endNum) {
		Row targetRow = sheet.createRow(endNum);
		Iterator<Cell> itCell = row.cellIterator();
		while(itCell.hasNext()){
			Cell cell = itCell.next();
			Cell newCell = targetRow.createCell(cell.getColumnIndex());
			newCell.setCellStyle(cell.getCellStyle());
			newCell.setCellValue(cell.getRichStringCellValue());			
		}
		return targetRow;
		
	}

	

}
