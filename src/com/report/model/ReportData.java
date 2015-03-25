package com.report.model;

import java.util.List;
import java.util.Map;

public class ReportData {
    private String date;
    private String programname;
    private List<List<String>> header;
    private List<Map<String, Object>> body;

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getProgramname() {
        return programname;
    }

    public void setProgramname(String programname) {
        this.programname = programname;
    }

    public List<List<String>> getHeader() {
        return header;
    }

    public void setHeader(List<List<String>> header) {
        this.header = header;
    }

    public List<Map<String, Object>> getBody() {
        return body;
    }

    public void setBody(List<Map<String, Object>> body) {
        this.body = body;
    }


}
