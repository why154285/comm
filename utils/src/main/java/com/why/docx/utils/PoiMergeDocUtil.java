package com.why.docx.utils;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author why
 */
public class PoiMergeDocUtil {
    public static void main(String[] args) {
        String[] srcDocxs = {"d:\\sampleTemplate\\1.docx", "d:\\sampleTemplate\\2.docx"};

        String destDocx = "d:\\sampleTemplate\\waz.docx";
        mergeDoc(srcDocxs, destDocx);
    }

    /**
     * 合并decx文件
     *
     * @param srcDocxs 需要合并的目标docx文件
     * @param destDocx 合并后的docx输出文件
     */
    public static void mergeDoc(String[] srcDocxs, String destDocx) {
        OutputStream dest = null;
        List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
        int length = null == srcDocxs ? 0 : srcDocxs.length;
        for (int i = 0; i < length; i++) {
            String doc = srcDocxs[i];
            OPCPackage srcPackage = null;
            try {
                srcPackage = OPCPackage.open(doc);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (null != srcPackage) {
                opcpList.add(srcPackage);
            }
        }
        int opcpSize = opcpList.size();
        if (opcpSize > 0) {
            try {
                dest = new FileOutputStream(destDocx);
                XWPFDocument srclDocument = new XWPFDocument(opcpList.get(0));
                CTBody srclBody = srclDocument.getDocument().getBody();
                if (opcpSize > 1) {
                    for (int i = 1; i < opcpSize; i++) {
                        OPCPackage src2Package = opcpList.get(i);
                        XWPFDocument src2Document = new XWPFDocument(src2Package);
                        CTBody src2Body = src2Document.getDocument().getBody();
                        appendBody(srclBody, src2Body);
                    }
                }
                srclDocument.write(dest);
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                IOUtils.closeQuietly(dest);
                System.gc();
            }
        }
    }

    /**
     * 合并文档内容
     *
     * @param src    目标文档
     * @param append 要合并的文档
     * @throws Exception 异常
     */
    private static void appendBody(CTBody src, CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src.set(makeBody);
    }
}
