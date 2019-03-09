package com.huitong.merge;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;


/**
 * <p></p>
 * author pczhao  <br/>
 * date  2019-03-09 16:58
 */

public class Docx2Pdf {
    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String docx = "D:\\logs\\1.docx";
        String pdf = "D:\\logs\\1.pdf";
        // 直接转换
        InputStream docxStream = new FileInputStream(docx);
        byte[] pdfData = docxToPdf(docxStream);
        com.sse.demo2.service.util.FileUtil.writeByte2File(pdf, pdfData);
        System.out.println("finished.");
    }

    /**
     * docx转成pdf
     *
     * @param docxStream docx文件流
     * @return 返回pdf数据
     * @throws Exception
     */
    public static byte[] docxToPdf(InputStream docxStream) throws Exception {
        ByteArrayOutputStream targetStream = null;
        XWPFDocument doc;
        try {
            doc = new XWPFDocument(docxStream);
            PdfOptions options = PdfOptions.create();
            targetStream = new ByteArrayOutputStream();
            PdfConverter.getInstance().convert(doc, targetStream, options);
            return targetStream.toByteArray();
        } catch (IOException e) {
            throw new Exception(e);
        } finally {
            targetStream.close();
        }
    }


}
