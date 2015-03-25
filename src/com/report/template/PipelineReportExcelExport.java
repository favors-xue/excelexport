package com.report.template;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
 

public class PipelineReportExcelExport extends AbstractExcelExport {

	private static short BODY_HEIGHT = 310;
	private static short TITLE_HEIGHT = 800;

	@Override
	public void buildHeader(Sheet sheet) {
		setHeaderStyle(sheet);
		createTitle(sheet); 
		createSubTitle(sheet);
		createCountTable(sheet);
	}
	
	@Override
	public void buildBody(Sheet sheet) {
		createTable(sheet);
		
	}
	
	@Override
	public void buildTail(Sheet sheet) {
		// TODO Auto-generated method stub
		
	}

	private void createTable(Sheet sheet) {
		buildTableFrame(sheet);
		initTableData(sheet);
		
	}

	private void initTableData(Sheet sheet) {
		initTableHeaderData(sheet);
		initTableBodyData(sheet);
		
	}

	private void initTableBodyData(Sheet sheet) {
		List<Map<String,Object>> result = getReportData().getBody();
		Iterator<Map<String,Object>> list = result.iterator();
		int index = 0;
		while(list.hasNext()){
			Map<String,Object> detail = (HashMap<String,Object>)list.next();
			Row row = sheet.getRow(15+index);
			Iterator<Cell> itCell = row.cellIterator();
			
			for(String key : detail.keySet()){
				itCell.next().setCellValue(detail.get(key).toString());
			}
		}
		
	}

	/**
	 * Build the Header of the Table
	 * @param sheet
	 */
	private void initTableHeaderData(Sheet sheet) {
		//Header
		Row row = sheet.getRow(13);
		Cell cell = row.getCell(0);
		cell.setCellValue("Client Information");
		cell = row.getCell(2);
		cell.setCellValue("Application Information");
		cell = row.getCell(5);
		cell.setCellValue("Loan Requested");
		cell = row.getCell(8);
		cell.setCellValue("Credit Scoring Offer");
		cell = row.getCell(12);
		cell.setCellValue("Sales Manager");
		
		//SubHeader
		row = sheet.getRow(14);
		cell = row.getCell(0);
		cell.setCellValue("Client ID");
		cell = row.getCell(1);
		cell.setCellValue("Legal Name");
		cell = row.getCell(2);
		cell.setCellValue("App ID");
		cell = row.getCell(3);
		cell.setCellValue("Current Stage");
		cell = row.getCell(4);
		cell.setCellValue("Status as of Date");
		cell = row.getCell(5);
		cell.setCellValue("Request Date");
		cell = row.getCell(6);
		cell.setCellValue("Loan Amount");
		cell = row.getCell(7);
		cell.setCellValue("Loan Term");
		cell = row.getCell(8);
		cell.setCellValue("Approval Date");
		cell = row.getCell(9);
		cell.setCellValue("CS%");
		cell = row.getCell(10);
		cell.setCellValue("CS Grading");
		cell = row.getCell(11);
		cell.setCellValue("Final Loan Amount");
		cell = row.getCell(12);
		cell.setCellValue("Sales Manager");
		
		
	}

	/**
	 * Build the Table Frame including Rows,Cells,Border_width,Foregroud-Color
	 * @param sheet
	 */
	private void buildTableFrame(Sheet sheet) {
		buildTableHeaderFrame(sheet);
		buildTableBodyFrame(sheet);
	}

	private void buildTableBodyFrame(Sheet sheet) {
		int rowNum = getReportData().getBody().size();
		for(int i=15;i<(15+rowNum);i++){
			Row row = sheet.createRow(i);
			for(int j=0;j<13;j++){
				Cell cell = row.createCell(j);
				setCellStyle(cell,j);
			}
		}
		
	}

	private void setCellStyle(Cell cell,int index) {
		switch(index){
			case 4: 
			case 5:
			case 8:cell.setCellStyle(getStyle().get("datecell"));break;
			case 11:
			case 6:cell.setCellStyle(getStyle().get("moneycell"));break;
			default:cell.setCellStyle(getStyle().get("commoncell"));break;
			
		}
		
	}

	private void buildTableHeaderFrame(Sheet sheet) {
		Row headerRow = sheet.createRow(13);
		int index = 1;
		for(int i=0;i<13;i++){
			headerRow.createCell(i);
			
			if(i==2||i==5||i==8||i==12){
				index++;
			}
				
			Cell cell = headerRow.getCell(i);
			cell.setCellStyle(getStyle().get("tableheader"+index));
		}
		headerRow.setHeight(BODY_HEIGHT);
		sheet.addMergedRegion(CellRangeAddress.valueOf("$A$14:$B$14"));
		sheet.addMergedRegion(CellRangeAddress.valueOf("$C$14:$E$14"));
		sheet.addMergedRegion(CellRangeAddress.valueOf("$F$14:$H$14"));
		sheet.addMergedRegion(CellRangeAddress.valueOf("$I$14:$L$14"));
		headerRow = sheet.createRow(14);
		headerRow.setHeight(BODY_HEIGHT);
		index = 1;
		for(int i=0;i<13;i++){
			headerRow.createCell(i);
			if(i==2||i==5||i==8||i==12){
				index++;
			}
				
			Cell cell = headerRow.getCell(i);
			cell.setCellStyle(getStyle().get("tableheader"+index));
		}
	}

	


	/**
	 * Build the count table in Header Field
	 * @param sheet
	 */
	private void createCountTable(Sheet sheet) {
		buildHeadTableFrame(sheet);
		initHeadTableData(sheet);
		
	}

	private void initHeadTableData(Sheet sheet) {
		List<List<String>> headerlist = getReportData().getHeader();
		for(int i=6;i<11;i++){
			Row row = sheet.getRow(i);
			List<String> detail = headerlist.get(i-6);
			Iterator<String> value = detail.iterator();
			for(int j=0;j<6;j++){
				
				if(j==1||j==4){
					j++;//�ϲ��ĵ�Ԫ��
				}					
				Cell cell = row.getCell(j);
				if(cell==null){
					break;
				}
				if(value.hasNext()){
					String v =  value.next();
					cell.setCellValue(v);
				}
			}				 
		} 
	}

	private void buildHeadTableFrame(Sheet sheet) {
		Row tableHeaderRow = sheet.createRow(6);
		tableHeaderRow.setHeight(BODY_HEIGHT);
		for(int i=0;i<6;i++){
			tableHeaderRow.createCell(i).setCellStyle(getStyle().get("greenheader"));			
		}
		sheet.addMergedRegion(CellRangeAddress.valueOf("$A$7:$B$7"));
		sheet.addMergedRegion(CellRangeAddress.valueOf("$D$7:$E$7"));
		for(int j=7;j<11;j++){
			Row tableRow = sheet.createRow(j);
			tableRow.setHeight(BODY_HEIGHT);
			for(int m=0;m<6;m++){
				if(j==10&&m==3){
					break;
				}
				tableRow.createCell(m).setCellStyle(getStyle().get("boldtablecell"));
			}
			int n =j+1;
			sheet.addMergedRegion(CellRangeAddress.valueOf("$A$"+n+":$B$"+n));
			if(j!=10){
				sheet.addMergedRegion(CellRangeAddress.valueOf("$D$"+n+":$E$"+n));
			}
		}
	}

	private void setHeaderStyle(Sheet sheet) {
		for(int i=0;i<6;i++){
			Row headerRow = sheet.createRow(i);	 
			for(int j=0;j<13;j++){
				headerRow.createCell(j).setCellStyle(getStyle().get("header"));
			}
		}		
	}

	private void createSubTitle(Sheet sheet) {
		Row subTitleRow = sheet.getRow(3);
		Cell cell = subTitleRow.getCell(1);
		//Program Name field
		cell.setCellValue("Program Name:");
		cell.setCellStyle(getStyle().get("subtitle"));
		cell = subTitleRow.getCell(2);		
		cell.setCellStyle(getStyle().get("subtitleblank"));
		cell.setCellValue(getReportData().getProgramname());
			
		//Generation Data field
		cell = subTitleRow.getCell(5);
		cell.setCellStyle(getStyle().get("subtitle"));
		cell.setCellValue("Generation Date:");		
		cell = subTitleRow.getCell(6);
		cell.setCellStyle(getStyle().get("dateblank"));
		cell.setCellValue(getReportData().getDate());
		
	}
	
	

	private void createTitle(Sheet sheet) {
		Row titleRow = sheet.getRow(1);	
		titleRow.setHeight(TITLE_HEIGHT);
		Cell cell = titleRow.getCell(6);		
		cell.setCellValue("Pipeline Report");	
		cell.setCellStyle(getStyle().get("title"));
		sheet.addMergedRegion(CellRangeAddress.valueOf("$G$2:$H$2"));
	}

	@Override
	public void buildFrame(Sheet sheet) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void setColumnWidth(Sheet sheet) {
		sheet.setColumnWidth(1, 14*256);
		sheet.setColumnWidth(2, 18*256);
		sheet.setColumnWidth(3, 13*256);
		sheet.setColumnWidth(4, 17*256);
		sheet.setColumnWidth(5, 14*256);
		sheet.setColumnWidth(6, 14*256);
		sheet.setColumnWidth(7, 14*256);
		sheet.setColumnWidth(8, 11*256);
		sheet.setColumnWidth(9, 13*256);
		sheet.setColumnWidth(10, 13*256);
		sheet.setColumnWidth(11, 13*256);
		sheet.setColumnWidth(12, 17*256);
		sheet.setColumnWidth(13, 25*256);
	}

	
	 
	
}
