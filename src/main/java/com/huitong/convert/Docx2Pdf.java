package com.huitong.convert;

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
        /*long startTime = System.currentTimeMillis();
        String docx = "D:\\logs\\1.docx";
        String pdf = "D:\\logs\\1.pdf";
        // 直接转换
        InputStream docxStream = new FileInputStream(docx);
        byte[] pdfData = docxToPdfByteArray(docxStream);
//        FileUtilHelper.writeByte2File(pdf, pdfData);
        System.out.println(getPdfDocumentPageNumber(new ByteArrayInputStream(pdfData), pdf));

        System.out.println("cost time:" + (System.currentTimeMillis() - startTime));
        System.out.println("finished.");*/
    }

    /**
     * docx转成pdf
     *
     * @param docxStream docx文件流
     * @return 返回pdf数据
     * @throws Exception
     */
    /*public static byte[] docxToPdfByteArray(InputStream docxStream) throws Exception {
        ByteArrayOutputStream targetStream;
        XWPFDocument doc;
        doc = new XWPFDocument(docxStream);
        PdfOptions options = PdfOptions.create();
        targetStream = new ByteArrayOutputStream();
        PdfConverter.getInstance().convert(doc, targetStream, options);
        targetStream.close();
        return targetStream.toByteArray();
    }

    private static int getPdfDocumentPageNumber(String pdfFilename) throws IOException, PDFException, PDFSecurityException {
        Document document = new Document();
        document.setFile(pdfFilename);
        return document.getPageTree().getNumberOfPages();
    }

    private static int getPdfDocumentPageNumber(InputStream pdfInStream, String pdfFilename) throws IOException, PDFException, PDFSecurityException {
        Document document = new Document();
        document.setInputStream(pdfInStream, pdfFilename);
        return document.getPageTree().getNumberOfPages();
    }*/


}
