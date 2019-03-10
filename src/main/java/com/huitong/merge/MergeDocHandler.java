package com.huitong.merge;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author pczhao
 * @email
 * @date 2018-10-10 15:06
 */

public class MergeDocHandler {
    private InputStream emptyDocxFilename = MergeDocHandler.class.getResourceAsStream("/doc/emptyxml.docx");

    static {
        ZipSecureFile.setMinInflateRatio(0.00001);
    }

    public MergeDocHandler() {
    }

    public static void main(String[] args) {
        try {
            testIMerge();
        } catch (XmlException e) {
            e.printStackTrace();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
//        testMerge();
//        testMergeDoc();
    }

    private static void testIMerge() throws XmlException, ParserConfigurationException, InvalidFormatException, SAXException, IOException {
        String inFile1 = Class.class.getClass().getResource("/").getPath() + "doc/1.docx";
        String inFile2 = Class.class.getClass().getResource("/").getPath() + "doc/2.docx";
        String inFile3 = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/dom1.doc";
        String inFile4 = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/dom2.doc";
        String outfile = Class.class.getClass().getResource("/").getPath() + "doc/0out.docx";
        MergeDocHandler handler = new MergeDocHandler();
        List<String> docs = new ArrayList<>();
        docs.add(inFile1);
        docs.add(inFile1);
        docs.add(inFile1);
        docs.add(inFile2);
        System.out.println(handler.iMergeMixedFileList2Docx(outfile, docs));
    }

    private static void testMerge() {
        String inFile1 = Class.class.getClass().getResource("/").getPath() + "doc/1.1.docx";
        String inFile2 = Class.class.getClass().getResource("/").getPath() + "doc/2.1.docx";
        String inFile3 = Class.class.getClass().getResource("/").getPath() + "doc/3.1.docx";
        String outfile = Class.class.getClass().getResource("/").getPath() + "doc/0out.docx";
        MergeDocHandler handler = new MergeDocHandler();
        List<String> docs = new ArrayList<>();
        docs.add(inFile1);
        docs.add(inFile2);
        docs.add(inFile3);
        handler.mergeDocx(outfile, docs);
    }

    private static void testMergeDoc() throws XmlException, ParserConfigurationException, InvalidFormatException, SAXException, IOException {
        String inFile1 = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/domparc1.doc";
        String inFile2 = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/domparc2.doc";
        String inFile3 = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/domimg.doc";
        String outfile = MergeDocHandler.class.getClassLoader().getResource("").getPath() + "doc/0out.docx";
        MergeDocHandler handler = new MergeDocHandler();
        List<String> docFiles = new ArrayList<>();
        docFiles.add(inFile1);
        docFiles.add(inFile2);
        docFiles.add(inFile3);
        handler.mergeDocList2Docx(outfile, docFiles);
    }

    /**
     * 将 doc 或 docx 文件列表 mergeList 合并到目标文件 destFilename，目标文件 destFilename 后缀是 .docx
     * doc 文件是生成的通知单，其他格式的 doc 不支持。
     *
     * @param destFilename
     * @param mergeList
     * @return
     */
    public boolean iMergeMixedFileList2Docx(final String destFilename, List<String> mergeList) throws InvalidFormatException, XmlException, IOException, ParserConfigurationException, SAXException {
        boolean result = false;
        if (!endsWithDocx(destFilename)) {
            return result;
        }
        com.sse.demo2.service.util.FileUtilHelper.deleteFile(destFilename);
        if (mergeList == null || mergeList.isEmpty()) {
            return result;
        }
        for (String filename : mergeList) {
            if (endsWithDocx(filename)) {
                if (!mergeDocx2Docx(destFilename, filename)) {
                    return result;
                }
            } else {
                if (!mergeDoc2Docx(destFilename, filename)) {
                    return result;
                }
            }
        }
        result = true;
        return result;
    }

    /**
     * 将所有 xml 文档列表合并到 docx 目标文件中
     *
     * @param destFilename 文档后缀是 .docx
     * @param mergeXmls    里面的文档后缀是 .xml，格式是 word xml不是 word2003xml
     * @return
     */
    public boolean mergeDocList2Docx(final String destFilename, List<String> mergeXmls) throws IOException, XmlException, ParserConfigurationException, InvalidFormatException, SAXException {
        boolean result = false;
        if (destFilename == null || mergeXmls == null) {
            return result;
        }
        Iterator<String> it = mergeXmls.iterator();
        while (it.hasNext()) {
            String next = it.next();
            if (next == null || !next.endsWith(".doc")) {
                it.remove();
            }
        }
        if (mergeXmls.size() == 0) {
            return result;
        }
        XWPFDocument desDocument = null;
        desDocument = new XWPFDocument(emptyDocxFilename);
        for (String filename : mergeXmls) {
            if (desDocument == null) {
                desDocument = new XWPFDocument(new FileInputStream(destFilename));
                XWPFRun run = desDocument.getLastParagraph().createRun();
                run.addBreak(BreakType.PAGE);
            }
            appendBody(desDocument, filename);
            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(destFilename));
            desDocument.write(bufferedOutputStream);
            bufferedOutputStream.close();
            desDocument = null;
        }
        result = true;
        return result;
    }

    /**
     * 将 doc 文件添加到 docx 文件后
     *
     * @param destFilename 目标文件，docx 格式
     * @param docFilename  待合并文件， doc 格式
     * @return
     */
    private boolean mergeDoc2Docx(final String destFilename, final String docFilename) throws IOException, XmlException, ParserConfigurationException, InvalidFormatException, SAXException {

        boolean result = false;
        XWPFDocument desDocument = null;
        if (!new File(destFilename).exists()) {
            /** 目标文件不存在。将 xml 文件转换成 docx 文件追加到空的 docx 文件中 */
            desDocument = new XWPFDocument(emptyDocxFilename);
        } else {
            /** 目标文件存在。 */
            desDocument = new XWPFDocument(new FileInputStream(destFilename));
            XWPFRun run = desDocument.getLastParagraph().createRun();
            run.addBreak(BreakType.PAGE);
        }
        appendBody(desDocument, docFilename);
        BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(destFilename));
        desDocument.write(bufferedOutputStream);
        desDocument.close();
        bufferedOutputStream.close();
        result = true;
        return result;
    }

    /**
     * 将 appendXmlFilename 文档的内容添加到 DesDocument 文档中
     *
     * @param desDocument
     * @param appendXmlFilename
     * @throws XmlException
     */
    private void appendBody(XWPFDocument desDocument, String appendXmlFilename) throws XmlException, InvalidFormatException, ParserConfigurationException, SAXException, IOException {
        XmlDomHandler domHandler = new XmlDomHandler(appendXmlFilename);
        List<XmlPicture> xmlAllPictures = domHandler.getXmlAllPictures();
        // 记录图片合并前及合并后的ID
        Map<String, String> map = new HashMap<>();
        for (XmlPicture picture : xmlAllPictures) {
            String before = picture.getRelationId();
            String after = desDocument.addPictureData(picture.getData(), picture.getType());
            if (before.equals(after)) {
                continue;
            }
            map.put(before, after);
        }
        XmlBody appendBody = domHandler.getXmlBody();
        if (appendBody == null) {
            return;
        }
        String appendBodyStr = appendBody.getCTBodyStr();
        appendBodyStr = convertImgIdString(appendBodyStr, map);
        CTBody desBody = desDocument.getDocument().getBody();
        appendBody(desBody, appendBodyStr);
    }

    /**
     * 将 mergeDocs 所有文档合并到 destFilename 文档中
     * 合并所有 docx 文件列表到 docx 目标文件中
     *
     * @param destFilename
     * @param mergeDocs
     * @return
     */
    public boolean mergeDocx(String destFilename, List<String> mergeDocs) {
        boolean result = false;
        if (destFilename == null || mergeDocs == null) {
            return result;
        }
        Iterator<String> it = mergeDocs.iterator();
        while (it.hasNext()) {
            String next = it.next();
            if (!next.toLowerCase().endsWith(".docx")) {
                it.remove();
            }
        }
        if (mergeDocs.size() == 0) {
            return result;
        }
        XWPFDocument desDocument = null;
        try {
            desDocument = new XWPFDocument(new FileInputStream(mergeDocs.remove(0)));
            for (String filename : mergeDocs) {
                if (desDocument == null) {
                    desDocument = new XWPFDocument(new FileInputStream(destFilename));
                }
                XWPFDocument appendDocument = new XWPFDocument(new FileInputStream(filename));
                appendBody(desDocument, appendDocument, true);
                BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(destFilename));
                desDocument.write(bufferedOutputStream);
                bufferedOutputStream.close();
                desDocument = null;
            }
            result = true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (XmlException e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 将 docx 源文件追加到目标文件中，目标文件后缀是 .docx
     *
     * @param destFilename   目标文件
     * @param appendFilename 源文件
     * @return
     */
    private boolean mergeDocx2Docx(final String destFilename, final String appendFilename) throws IOException, InvalidFormatException, XmlException {
        boolean result = false;
        /** 目标文件不存在 */
        if (!new File(destFilename).exists()) {
            return com.sse.demo2.service.util.FileUtilHelper.copyFile(destFilename, appendFilename);
        }
        XWPFDocument desDocument = new XWPFDocument(new FileInputStream(destFilename));
        XWPFDocument srcDocument = new XWPFDocument(new FileInputStream(appendFilename));
        appendBody(desDocument, srcDocument, true);
        BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(destFilename));
        desDocument.write(bufferedOutputStream);
        bufferedOutputStream.close();
        result = true;
        return result;
    }

    public static boolean endsWithDoc(final String filename) {
        if (filename == null || !filename.trim().toLowerCase().endsWith(".doc")) {
            return false;
        }
        return true;
    }

    public static boolean endsWithDocx(final String filename) {
        if (filename == null || !filename.trim().toLowerCase().endsWith(".docx")) {
            return false;
        }
        return true;
    }

    public static boolean endsWithXml(final String filename) {
        if (filename == null || !filename.trim().toLowerCase().endsWith(".xml")) {
            return false;
        }
        return true;
    }

    /**
     * 将 append Document 追加到 des Document 文档中
     *
     * @param des
     * @param append
     * @throws InvalidFormatException
     * @throws XmlException
     */
    private static void appendBody(XWPFDocument des, XWPFDocument append, final boolean appendBreak) throws InvalidFormatException, XmlException {
        if (appendBreak) {
            XWPFRun run = des.getLastParagraph().createRun();
            run.addBreak(BreakType.PAGE);
        }
        CTBody desBody = des.getDocument().getBody();
        CTBody appendBody = append.getDocument().getBody();
        List<XWPFPictureData> allPictures = append.getAllPictures();
        // 记录图片合并前及合并后的ID
        Map<String, String> map = new HashMap<>();
        for (XWPFPictureData picture : allPictures) {
            String before = append.getRelationId(picture);
            //将原文档中的图片加入到目标文档中
            String after = null;
            after = des.addPictureData(picture.getData(), picture.getPictureType());
            map.put(before, after);
        }
        convertImgIds(appendBody, map);
        appendBody(desBody, appendBody);
    }

    /**
     * 将文档内容（CTBody 格式）中所有图片的老 relationID 替换成新文档中的 relationID
     *
     * @param body
     * @param map
     */
    private static void convertImgIds(CTBody body, Map<String, String> map) {
        String bodyText = body.xmlText();
        if (map == null || map.size() == 0) {
            return;
        }
        bodyText = convertImgIdString(bodyText, map);
        CTBody replacedBodyBody = null;
        try {
            replacedBodyBody = CTBody.Factory.parse(bodyText);
        } catch (XmlException e) {
            e.printStackTrace();
        }
        body.set(replacedBodyBody);
    }

    /**
     * 将文档内容（String格式）中所有图片的老 relationID 替换成新文档中的 relationID
     *
     * @param content
     * @param map
     * @return
     */
    private static String convertImgIdString(String content, Map<String, String> map) {
        StringBuilder sb = new StringBuilder();
        Matcher matcher = Pattern.compile("<a:blip.*?>").matcher(content);
        int lastStart = 0;
        for (int i = 0; i < map.size() && matcher.find(); i++) {
            int start = matcher.start();
            String img = matcher.group();
            sb.append(content.substring(lastStart, start));
            sb.append(replaceImgId(img, map));
            lastStart = start + img.length();
        }
        if (lastStart < content.length()) {
            sb.append(content.substring(lastStart));
        }
        return sb.toString();
    }

    /**
     * 将 body 中的图片内容 的图片ID进行替换，原始的和新文档中的最新 id 替换
     *
     * @param imgContent
     * @param map
     * @return
     */
    private static String replaceImgId(String imgContent, Map<String, String> map) {
        for (Map.Entry<String, String> entry : map.entrySet()) {
            if (imgContent.contains(entry.getKey())) {
                imgContent = imgContent.replace(entry.getKey(), entry.getValue());
                break;
            }
        }
        return imgContent;
    }

    /**
     * 将 append 的 Body 内容添加到 des Body 中
     *
     * @param des
     * @param append
     * @throws XmlException
     */
    private static void appendBody(CTBody des, CTBody append) throws XmlException {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = des.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        //将两个文档的xml内容进行拼接
        StringBuilder sb = new StringBuilder();
        sb.append(prefix);
        sb.append(mainPart);
        sb.append(addPart);
        sb.append(sufix);
        CTBody makeBody = CTBody.Factory.parse(sb.toString());
        des.set(makeBody);
    }

    /**
     * 专用于 xml 格式的文档内容添加到 docx 文档 body 中
     *
     * @param des
     * @param appendString
     * @throws XmlException
     */
    private static void appendBody(CTBody des, String appendString) throws XmlException {
        String srcString = des.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        //将两个文档的xml内容进行拼接
        StringBuilder sb = new StringBuilder();
        sb.append(prefix);
        sb.append(mainPart);
        sb.append(addPart);
        sb.append(sufix);
        CTBody makeBody = CTBody.Factory.parse(sb.toString());
        des.set(makeBody);
    }

}
