package com.why.docx.model;


import java.util.List;
import java.util.Map;

/**
 * @author why
 */
public class DocxModel {
    private Map<String,String> textMap;
    private List<List<String[]>> tableList;

    public Map<String, String> getTextMap() {
        return textMap;
    }

    public void setTextMap(Map<String, String> textMap) {
        this.textMap = textMap;
    }

    public List<List<String[]>> getTableList() {
        return tableList;
    }

    public void setTableList(List<List<String[]>> tableList) {
        this.tableList = tableList;
    }
}
