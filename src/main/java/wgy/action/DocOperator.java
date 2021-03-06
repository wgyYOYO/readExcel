package wgy.action;

import org.apache.commons.compress.utils.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import wgy.entity.CertificateData;
import wgy.entity.CertificateField;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;

public class DocOperator {


    public static File toCumstomDoc(String templatePath, String outputPath, CertificateData certificateData)
            throws IOException, XmlException {
        XWPFDocument document = new XWPFDocument(new FileInputStream(templatePath));


        Set<String> keys = certificateData.getKeys();


        for (XWPFParagraph paragraph : document.getParagraphs()) {
            XmlCursor cursor = paragraph.getCTP().newCursor();
            cursor.selectPath(
                    "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");


            List<XmlObject> xmlObjects = toXmlObjects(cursor);


            for (XmlObject xmlObject : xmlObjects) {
                CTR ctr = CTR.Factory.parse(xmlObject.xmlText());
                XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
                String text = bufferrun.getText(0);
                String conformingKey = containsKey(text, keys);
                if (conformingKey != null) {
                    CertificateField fieldInfo = certificateData.getValue(conformingKey);
                    text = text.replace(toTemplateKey(conformingKey), fieldInfo.getContent());
                    bufferrun.setFontSize(fieldInfo.getFontSize());
                    bufferrun.setFontFamily(fieldInfo.getFontFamily());
                    bufferrun.setText(text, 0);
                    bufferrun.setBold(fieldInfo.getIsBold());
                }


                xmlObject.set(bufferrun.getCTR());
            }
        }


        FileOutputStream out = new FileOutputStream(outputPath);
        document.write(out);


        out.close();
        document.close();


        return new File(outputPath);
    }


    public static List<XmlObject> toXmlObjects(XmlCursor docXmlCursor) {
        List<XmlObject> xmlObjects = Lists.newArrayList();


        while (docXmlCursor.hasNextSelection()) {
            docXmlCursor.toNextSelection();
            xmlObjects.add(docXmlCursor.getObject());
        }


        return xmlObjects;
    }


    public static String containsKey(String text, Set<String> keys) {
        String conforming = null;


        if (StringUtils.isEmpty(text)) {
            return conforming;
        }


        for (String key : keys) {
            if (text.contains(key)) {
                conforming = key;
                break;
            }
        }


        return conforming;
    }


    public static String toTemplateKey(String key) {
        if (StringUtils.isEmpty(key)) {
            return null;
        }


        return "${" + key + "}";
    }


    public static String addBlankSpace(String text) {
        StringBuffer sb = new StringBuffer();


        if (text == null) {
            return null;
        }


        char[] chars = text.toCharArray();


        String regex = "[\u4E00-\u9FA5]{1}";
        for (char aChar : chars) {
            String str = aChar + "";
            if (StringUtils.isBlank(str)) {
                continue;
            }


            sb.append(aChar);


            if (str.matches(regex)) {
                sb.append(" ");
            }
        }


        return sb.toString();
    }


    public static void main(String[] args) throws IOException, XmlException {
        String templatePath = "C:\\Users\\a1579\\Desktop\\template.docx";
        String outputPath = "C:\\Users\\a1579\\Desktop\\custom.docx";
        CertificateData data = new CertificateData();
        data.put(new CertificateField("?????????", addBlankSpace("??????"), 34));
        data.put(new CertificateField("??????????????????", "???????????????\"??????????????????????????????\"??????????????????????????????", 18));
        data.put(new CertificateField("??????????????????",
                "Congratulations on completion of training program of \"Helicopter Pilot Qualification\"", 15));
        data.put(new CertificateField("????????????", "100224512", 18));
        data.put(new CertificateField("????????????", addBlankSpace("??????"), 25));
        data.put(new CertificateField("????????????", "2020???3???21???", 15));
        toCumstomDoc(templatePath, outputPath, data);
    }
}