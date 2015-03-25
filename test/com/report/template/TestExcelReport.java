package com.report.template;

import java.sql.Date;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.report.model.ReportData;
 

public class TestExcelReport {

	public static void main(String[] args)
	{
		 
		AbstractExcelExport a = new PipelineReportExcelExport();
		ReportData h = new ReportData();
		List<List<String>> result = new ArrayList<List<String>>();
		List<String> details = new ArrayList<String>();
		
		details.add("Active");
		details.add("153");
		details.add("Inactive");
		details.add("72"); 
		result.add(details);
		details = new ArrayList<String>();
		details.add("Pending Verification");
		details.add("16");
		details.add("Withdrawn");
		details.add("16"); 
		result.add(details);
		details = new ArrayList<String>();
		details.add("Pending Review");
		details.add("25");
		details.add("Rejected");
		details.add("24"); 
		result.add(details);
		details = new ArrayList<String>();
		details.add("Pending Approval");
		details.add("34");
		details.add("Disbursed");
		details.add("32"); 
		result.add(details);
		details = new ArrayList<String>();
		details.add("Pending Confirmation");
		details.add("78"); 
		result.add(details);
		h.setHeader(result);
		h.setDate("2013-09-08");
		h.setProgramname("Program A");
		
		
		Map<String,Object> test = new HashMap<String,Object>();
		test.put("test1", 1231);
		test.put("test2", Date.valueOf("2008-08-09"));
		test.put("test4", 4000);
		List<Map<String,Object>> result1 = new ArrayList<Map<String,Object>>();
		for(int i=0;i<10;i++){
			result1.add(test);
		}
		h.setBody(result1);
		
		a.setReportData(h);
		a.ExcelExport("test6.xlsx");
		
	}
}
