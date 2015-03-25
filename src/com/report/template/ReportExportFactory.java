package com.report.template;

import com.report.model.ReportInfo;

public class ReportExportFactory {
	
	private static ReportExportFactory reportFactoryBean= null;
	
	private ReportExportFactory(){}
	
	public static ReportExportFactory getInstance(){
		if(reportFactoryBean!=null){
			return reportFactoryBean;
		}
		else{
			reportFactoryBean = new ReportExportFactory();
			return reportFactoryBean;
		}			
	}
	
	public AbstractReportExport createReportExport(ReportInfo reportInfo){
		AbstractReportExport reportBean = null;
		if(reportInfo.getTemplateName().equalsIgnoreCase("PipelineReport")){
			reportBean = new PipelineReportExport();
			reportBean.setReportInfo(reportInfo);			
		}
		return reportBean;
	}
	
}
