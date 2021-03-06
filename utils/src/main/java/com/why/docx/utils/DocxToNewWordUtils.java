package com.why.docx.utils;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author why
 */
public class DocxToNewWordUtils {
    /**
     * 根据模板生成新word文档
     * 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
     *
     * @param inputUrl  模板存放地址
     * @param outputUrl 输出文件地址
     * @param textMap   需要替换的信息集合
     * @param tableList 需要插入的表格信息集合
     * @return 成功返回true，失败返回false
     */
    public static boolean changeWord(String inputUrl, String outputUrl, Map<String, String> textMap, List<String[]> tableList) {
        // 模板转换默认成功
        boolean changeFlag = true;
        try {
            // 获取docx解析对象
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
            // 解析替换文本段落对象
            DocxToNewWordUtils.changeText(document, textMap);
            // 解析替换表格对象
            DocxToNewWordUtils.changeTable(document, textMap, tableList);

            //生成新的word
            File file = new File(outputUrl);
            FileOutputStream stream = new FileOutputStream(file);
            document.write(stream);
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
            changeFlag = false;
        }
        return changeFlag;
    }

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap) {
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落是否需要进行替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText(changeValue(run.toString(), textMap), 0);
                }
            }
        }
    }

    /**
     * 替换表格对象方法 -- 每个表格插入的内容相同
     *
     * @param document  docx解析对象
     * @param textMap   需要替换的信息集合
     * @param tableList 需要插入的表格信息集合
     */
    public static void changeTable(XWPFDocument document, Map<String, String> textMap, List<String[]> tableList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表格
            XWPFTable table = tables.get(i);
            changeOneTable(table, tableList, textMap);
        }

    }

    /**
     * 替换表格对象方法--每个表格插入的内容不同
     *
     * @param document         docx解析对象
     * @param textMap          需要替换的信息集合
     * @param tableContentList 需要插入的表格信息集合
     */
    public static void changeTableList(XWPFDocument document, Map<String, String> textMap, List<List<String[]>> tableContentList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0, contentSize = tableContentList.size(), tableSize = tables.size(); i < tableSize; i++) {
            List<String[]> tableList = new ArrayList<String[]>();
            if (i < contentSize) {
                tableList = tableContentList.get(i);
            }
            //只处理行数大于等于2的表格，且不循环表格
            XWPFTable table = tables.get(i);
            changeOneTable(table, tableList, textMap);
        }
    }

    /**
     * 改变一个表格的内容，将table的内容设置为tableList或者替换掉
     *
     * @param table     需要替换的表
     * @param tableList 表内容的集合
     * @param textMap   需要替换的占位符
     */
    public static void changeOneTable(XWPFTable table, List<String[]> tableList, Map<String, String> textMap) {
        if (table.getRows().size() > 1) {
            // 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
            if (checkText(table.getText())) {
                List<XWPFTableRow> rows = table.getRows();
                eachTable(rows, textMap);
            } else {
                if (tableList.size() < 1) {
                    return;
                }
                System.out.println("插入" + table.getText());
                insertTable(table, tableList);
            }
        }
    }

    /**
     * 遍历表格
     *
     * @param rows    表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows, Map<String, String> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap), 0);
                        }
                    }
                }
            }
        }
    }

    /**
     * 为表格插入数据，行数不够添加新行
     *
     * @param table     需要插入数据的表格
     * @param tableList 插入数据集合
     */
    public static void insertTable(XWPFTable table, List<String[]> tableList) {
        //创建行，根据需要插入的数据添加新行，不处理表头
        for (int i = 1; i < tableList.size(); i++) {
            //创建表格
            table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 1; i < rows.size(); i++) {
            XWPFTableRow newRow = table.getRow(i);
            List<XWPFTableCell> cells = newRow.getTableCells();
            for (int j = 0; j < cells.size(); j++) {
                XWPFTableCell cell = cells.get(j);
                if (i > tableList.size()) {
                    return;
                }
                cell.setText(tableList.get(i - 1)[j]);
            }
        }
    }

    /**
     * 判断文本中是否包含$
     *
     * @param text 文本
     * @return 包含返回true，不包含返回false
     */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.indexOf("$") != -1) {
            check = true;
        }
        return check;
    }

    /**
     * 匹配输入信息与模板
     *
     * @param value   模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap) {
        if (value.contains("${")) {
            value = value.replace("${", "").replaceAll("}", "").trim();
            if (textMap.containsKey(value)) {
                return textMap.get(value);
            }
            return "";
        } else {
            return value;
        }
    }

    public static void main(String[] args) {
        //模板文件地址
        String inputUrl = "d:\\sampleTemplate\\1.docx";
        //新模板地址
        String outputUrl = "d:\\sampleTemplate\\11test.docx";
        Map<String, String> testMap = new HashMap<String, String>();
        testMap.put("name", "小明");
        testMap.put("sex", "男");
        testMap.put("age", "18");
        testMap.put("txtWorkMode", "0000");
        testMap.put("address", "软件园");

        List<String[]> testList = new ArrayList<String[]>();
        testList.add(new String[]{"1","1AA","1BB","1CC"});
        testList.add(new String[]{"2","2AA","2BB","2CC"});
        testList.add(new String[]{"3","3AA","3BB","3CC"});
        testList.add(new String[]{"4","4AA","4BB","4CC"});
        DocxToNewWordUtils.changeWord(inputUrl, outputUrl, testMap, testList);
    }
}
