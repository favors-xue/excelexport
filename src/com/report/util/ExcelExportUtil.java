package com.report.util;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * 
 * @author Shawn¡¡Xue
 *
 */
public class ExcelExportUtil {
	
	/**
	 * cell styles used for formatter
	 * @param wb
	 * @return
	 */
    public static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
        CellStyle style = wb.createCellStyle(); 
        DataFormat formatter = wb.createDataFormat();
        
        //Font
        Font Arial_16_bold = wb.createFont();
        Arial_16_bold.setFontHeightInPoints((short)16);
        Arial_16_bold.setFontName("Arial");
        Arial_16_bold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        
        Font Arial_10_bold = wb.createFont();
        Arial_10_bold.setFontHeightInPoints((short)10);
        Arial_10_bold.setFontName("Arial");
        Arial_10_bold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        
        Font Arial_10_blue = wb.createFont();
        Arial_10_blue.setFontHeightInPoints((short)10);
        Arial_10_blue.setFontName("Arial");
        Arial_10_blue.setColor(IndexedColors.BLUE.getIndex());
        
        Font Arial_10 = wb.createFont();
        Arial_10.setFontHeightInPoints((short)10);
        Arial_10.setFontName("Arial"); 
        
        
        //Style
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("header", style);
        
        
        
        style = wb.createCellStyle();
        style.setFont(Arial_16_bold);         
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        styles.put("title", style);

         
        style = wb.createCellStyle();
        style.setFont(Arial_10_bold);
        styles.put("subtitle", style);
        
        style = wb.createCellStyle();
        style.setFont(Arial_10_bold); 
        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("subtitleblank", style);
        
        style = wb.createCellStyle();
        style.setFont(Arial_10_bold); 
        style.setBorderBottom(CellStyle.BORDER_MEDIUM);        
        style.setDataFormat(formatter.getFormat("dd/mm/yyyy"));
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("dateblank", style);
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        style.setFillForegroundColor(IndexedColors.LIME.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_bold);       
        styles.put("greenheader", style);
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10);
        styles.put("boldtablecell", style);
        
              
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(Arial_10_bold);       
        styles.put("tableheader1", style);
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_bold);       
        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(79, 129, 189)).getIndexed());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("tableheader2", style);
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_bold);       
        style.setFillForegroundColor(IndexedColors.TAN.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("tableheader3", style);
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_bold);       
        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(0, 176, 240)).getIndexed());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("tableheader4", style);
        
        
        style = wb.createCellStyle();
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_bold);       
        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 100, 162)).getIndexed());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("tableheader5", style);
        
        style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(Arial_10_blue);
        styles.put("commoncell", style);
        
        style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(formatter.getFormat("dd/mm/yyyy"));
        style.setFont(Arial_10_blue);
        styles.put("datecell", style);
        
        style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(formatter.getFormat("0"));
        style.setFont(Arial_10_blue);
        styles.put("numcell", style);
        
        style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderBottom(CellStyle.BORDER_THIN); 
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setDataFormat(formatter.getFormat("$#,##0"));
        style.setFont(Arial_10_blue);
        styles.put("moneycell", style);        

        return styles;
    }
}
