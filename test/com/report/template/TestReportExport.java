package com.report.template;

import java.sql.Date;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.report.model.ReportInfo;


public class TestReportExport {
	public static void main(String[] args){
		ReportInfo info = new ReportInfo();
		info.setInportFileName("reporttemplate\\PipelineReportTemplate.xlsx");
		info.setExportFileName("test123.xlsx");
		Map<String,Map<String,Object>> data = new HashMap<String,Map<String,Object>>();
		Map<String,Object> detail = new HashMap<String,Object>();
		detail.put("ProgramName", "Program A");
		detail.put("GenerationDate", Date.valueOf("2013-09-08"));
		detail.put("Active", Double.valueOf("153"));
		detail.put("Inactive", Double.valueOf("72"));
		detail.put("Verification", Double.valueOf("16"));
		detail.put("Withdrawn", Double.valueOf("16"));
		detail.put("Review", 25);
		detail.put("Rejected", 24);
		detail.put("Approval", 34);
		detail.put("Disbursed", 36);
		detail.put("Confirmation", 78);
		detail.put("Disbursed", 36);
		detail.put("Confirmation", 78); 
		
		List<Map<String,Object>> result = new ArrayList<Map<String,Object>>();
		for(int i=0;i<10;i++){
			Map<String,Object> resultInfo = new HashMap<String,Object>();
			resultInfo.put("sManager","Victor Lee");
			resultInfo.put("finalLAmount", 40000);
			resultInfo.put("cGrading", "Low Risk");
			resultInfo.put("CS", 75);
			resultInfo.put("aDate", Date.valueOf("2013-8-30"));
			resultInfo.put("lTerm", "180 days");
			resultInfo.put("lAmount", 50000);
			resultInfo.put("rDate", Date.valueOf("2013-8-1"));
			resultInfo.put("sDate", Date.valueOf("2013-8-30"));
			resultInfo.put("cStage", "Pending Confirmation");
			resultInfo.put("aID", "100008 001");
			resultInfo.put("lName", "AMP  Holdings");
			resultInfo.put("cID", "100008");
			result.add(resultInfo);
		}
		detail.put("List",result);
		data.put("refined 1", detail);
		
		info.setReportData(data);
		info.setTemplateName("PipelineReport");
		AbstractReportExport report = ReportExportFactory.getInstance().createReportExport(info);
		report.reportExport();
		
	}
}
