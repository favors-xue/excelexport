package com.report.template;
 
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.report.model.ReportData;
import com.report.util.ExcelExportUtil;

public abstract class AbstractExcelExport {

	private Map<String, CellStyle> style;
	private ReportData reportData;
	
	
	public ReportData getReportData() {
		return reportData;
	}

	public void setReportData(ReportData reportData) {
		this.reportData = reportData;
	}
	
	public Map<String,CellStyle> getStyle()
	{
		return this.style;
	}
	
	public void setStyle(Map<String,CellStyle> style)
	{
		this.style = style;
	}
	
	public void ExcelExport(String filename)
	{
		FileOutputStream out = null;
		Workbook wb = null;
		
				
		if(filename.indexOf(".xlsx")>0)
		{
			wb = new XSSFWorkbook();
		}
		else if(filename.indexOf(".xls")>0)
		{
			wb = new HSSFWorkbook();
		}
		else
		{
			return;
		}
			
		setStyle(ExcelExportUtil.createStyles(wb));
		Sheet sheet = wb.createSheet("Test");
		buildHeader(sheet);
		buildBody(sheet);
		buildTail(sheet);
		buildFrame(sheet);
		setColumnWidth(sheet);
		 
		try {
			out = new FileOutputStream(filename);
			wb.write(out);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        finally
        {
        	try {
        		if(out!=null)
        		{
        			out.close();
        		}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
	}
	
	 

	public abstract void buildHeader(Sheet sheet);
	
	public abstract void buildBody(Sheet sheet);
	
	public abstract void buildTail(Sheet sheet);
	
	public abstract void buildFrame(Sheet sheet);
	
	public abstract void setColumnWidth(Sheet sheet);
	
}
