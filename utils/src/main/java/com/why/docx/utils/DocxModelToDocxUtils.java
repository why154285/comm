package com.why.docx.utils;

import com.why.docx.model.DocxModel;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @author why
 */
public class DocxModelToDocxUtils {
    private static final String TEMP_PATH_1 = "d:\\sampleTemplate\\1.docx";
    private static final String TEMP_PATH_2 = "d:\\sampleTemplate\\2.docx";


    public static boolean changeWord(String inputUrl, String outputUrl, List<DocxModel> docxModelList){
        // 默认转换成功
        boolean changeFlag = true;
        Iterator<DocxModel> iterator = docxModelList.iterator();

        try {
            while (iterator.hasNext()){
                // 获取docx解析对象
                XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));
                DocxModel docxModel =iterator.next();
                Map<String,String> textMap = docxModel.getTextMap();
                List<List<String[]>> tableList = docxModel.getTableList();
                // 解析替换文本段落对象
                DocxToNewWordUtils.changeText(document,textMap);
                // 解析替换表格对象
                DocxToNewWordUtils.changeTableList(document,textMap,tableList);

                //替换后的文件存放在临时文件中
                File file = new File(TEMP_PATH_1);
                FileOutputStream stream = new FileOutputStream(file);
                document.write(stream);
                stream.close();
                // 替换模板后续处理

            }
        }catch (Exception e){
            e.printStackTrace();
            changeFlag = false;
        }
        return changeFlag;
    }

    private static void fillAfter(String outputUrl){
        File fileTarget = new File(outputUrl);
        if (fileTarget.exists()){
            // 若原文件存在，则先拷贝到临时文件中
            String[] targetPath = {outputUrl};
            PoiMergeDocUtil.mergeDoc(targetPath,TEMP_PATH_2);
            // 把模板文件和备份的原文件写入到指定的路径中
            String[] sourcePath = {TEMP_PATH_1,TEMP_PATH_2};
            PoiMergeDocUtil.mergeDoc(sourcePath,outputUrl);
        }else {
            String[] sourcePath = {TEMP_PATH_1};
            PoiMergeDocUtil.mergeDoc(sourcePath,outputUrl);
        }
        System.gc();
        System.out.println("====================");
        File file1 = new File(TEMP_PATH_1);
        if (file1.exists()){
            file1.delete();
        }
        File file2 = new File(TEMP_PATH_2);
        if (file2.exists()){
            file2.delete();
        }

    }

}
