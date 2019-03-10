package com.huitong.merge;

import lombok.Data;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author pczhao
 * @email
 * @date 2018-10-16 10:58
 */

@Data
public class XmlDomHandler {
    private Document doc;
    private String filename;

    public XmlDomHandler(String filename) throws IOException, SAXException, ParserConfigurationException {
        this.filename = filename;
        doc = loadXmlDocument(filename);
    }

    public static final Pattern MARK_PATTERN = Pattern.compile("\\$\\{(.*?)}");
    /**
     * ORI_OUT: true 表示原样输出， false 表示使用空串输出
     */
    public static boolean ORI_OUT = true;

    /**
     * 判断字符串是否含有 ${}
     *
     * @param content
     * @return
     */
    public static boolean matched(String content) {
        if (content == null) {
            return false;
        }
        Matcher matcher = MARK_PATTERN.matcher(content);
        if (matcher.find()) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 判断 xml 文件是否是 word xml 文件
     *
     * @param filename
     * @return
     */
    public static boolean isWordXml(final String filename) {
        String content = com.sse.demo2.service.util.FileUtilHelper.getFileContentStr(filename);
        return isWordXmlContent(content);
    }

    public static boolean isWordXml(final byte[] content) {
        String xmlFileContent = new String(content, Charset.forName("utf-8"));
        return isWordXmlContent(xmlFileContent);
    }

    private static boolean isWordXmlContent(final String content) {
        if (content.contains("<pkg:package")) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 提取占位符中的字符串,如果没有原样返回。
     *
     * @param text
     * @return
     */
    public static String getPlaceStr(String text) {
        String result = null;
        if (text == null) {
            return result;
        }
        Matcher matcher = MARK_PATTERN.matcher(text);
        if (matcher.find()) {
            result = matcher.group(1).trim();
        } else {
            result = text;
        }
        return result;
    }

    /**
     * 将字符串进行替换，如果在映射中不存在，原样返回。
     *
     * @param text
     * @param marks
     * @return
     */
    public static String replaceMatchedText(String text, Map<String, String> marks) {
        if (text == null) {
            return "";
        }
        while (true) {
            Matcher matcher = MARK_PATTERN.matcher(text);
            boolean replaced = false;
            while (matcher.find() && !replaced) {
                String findKey = matcher.group(1).trim();
                String allPlaceHolder = matcher.group();
                if (marks.containsKey(findKey)) {
                    text = replaceText(text, matcher.start(), allPlaceHolder.length(), marks.get(findKey));
                    replaced = true;
                } else if (!ORI_OUT) {
                    text = replaceText(text, matcher.start(), allPlaceHolder.length(), "");
                    replaced = true;
                }
            }
            if (!replaced) {
                break;
            }
        }
        return text;
    }

    /**
     * 将匹配的字符串使用 place 进行替换
     *
     * @param ori
     * @param start
     * @param length
     * @param place
     * @return
     */
    private static String replaceText(String ori, final int start,
                                      final int length, final String place) {
        StringBuilder sb = new StringBuilder();
        sb.append(ori.substring(0, start));
        sb.append(place);
        if (start + length < ori.length()) {
            sb.append(ori.substring(start + length));
        }
        return sb.toString();
    }

    /**
     * 载入 xml 文档
     *
     * @param filename
     * @return
     */
    public Document loadXmlDocument(String filename) throws ParserConfigurationException, IOException, SAXException {
        Document doc = null;
        doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(filename);
        doc.normalize();
        return doc;
    }

    public boolean flushFile(String outfilename) throws IOException {
        return flushFile(doc.getDocumentElement(), outfilename);
    }

    private boolean flushFile(Element document, String outfile) throws IOException {
        boolean result = false;
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        if (!flushByteArray(document, byteArrayOutputStream)) {
            return result;
        }
        BufferedOutputStream outputStream = null;
        outputStream = new BufferedOutputStream(new FileOutputStream(outfile));
        outputStream.write(byteArrayOutputStream.toByteArray());
        IOUtil.closeSilently(outputStream);
        result = true;
        return result;
    }

    private boolean flushByteArray(Element document, ByteArrayOutputStream outputStream) {
        boolean flag = false;
        if (document == null || outputStream == null) {
            return flag;
        }
        TransformerFactory transFactory = TransformerFactory.newInstance();
        Transformer transformer = null;
        try {
            transformer = transFactory.newTransformer();
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        }
        DOMSource source = new DOMSource();
        source.setNode(document);
        StreamResult result = new StreamResult();
        result.setOutputStream(outputStream);
        try {
            transformer.transform(source, result);
            flag = true;
        } catch (TransformerException e) {
            e.printStackTrace();
        }
        return flag;
    }

    /**
     * 将文档中所有 <w:t> 标签的内容进行占位符替换
     *
     * @param mark1
     * @return
     */
    private boolean replaceXmlParagraphContext(Map<String, String> mark1) {
        Element root = doc.getDocumentElement();
        return replaceWTContext(root, mark1);
    }

    /**
     * 对节点 root 及子节点所有 <w:t> 标签的元素进行替换
     *
     * @param tagMap
     * @return
     */
    private boolean replaceWTContext(Element root, Map<String, String> tagMap) {
        NodeList wtNodeList = root.getElementsByTagName("w:t");
        for (int i = 0; i < wtNodeList.getLength(); i++) {
            Node item = wtNodeList.item(i);
            String text;
            if (item.hasChildNodes() && matched(text = item.getFirstChild().getNodeValue())) {
                item.getFirstChild().setNodeValue(replaceMatchedText(text, tagMap));
            }
        }
        return true;
    }


    /**
     * 将文档中图片占位符进行替换
     *
     * @param marks
     */
    private boolean replaceImgData(Map<String, String> marks) {
        NodeList imgList = doc.getElementsByTagName("pkg:binaryData");
        for (int i = 0; i < imgList.getLength(); i++) {
            Element item = (Element) imgList.item(i);
            if (item.hasChildNodes() && matched(item.getFirstChild().getNodeValue())) {
                String key = getPlaceStr(item.getFirstChild().getNodeValue());
                item.getFirstChild().setNodeValue(marks.get(key));
            }
        }
        return true;
    }

    public boolean renderSimple(Map<String, Object> data) {
        boolean result = false;
        Map<String, String> strDataMap = new HashMap<>();
        Map<String, List> tbDataMap = new HashMap<>();
        for (Map.Entry<String, Object> entry : data.entrySet()) {
            if (entry.getValue() instanceof String) {
                strDataMap.put(entry.getKey(), (String) entry.getValue());
            } else if (entry.getValue() instanceof List) {
                tbDataMap.put(entry.getKey(), (List) entry.getValue());
            }
        }
        replaceTableSimple(tbDataMap);
        replaceXmlParagraphContext(strDataMap);
        replaceImgData(strDataMap);
        return result;
    }

    public boolean renderMap(Map<String, Object> data) {
        boolean result = false;
        Map<String, String> strDataMap = new HashMap<>();
        Map<String, List> tbDataMap = new HashMap<>();
        for (Map.Entry<String, Object> entry : data.entrySet()) {
            if (entry.getValue() instanceof String) {
                strDataMap.put(entry.getKey(), (String) entry.getValue());
            } else if (entry.getValue() instanceof List) {
                tbDataMap.put(entry.getKey(), (List) entry.getValue());
            }
        }
        if (!replaceTableMap(tbDataMap)) {
            return result;
        }
        if (!replaceXmlParagraphContext(strDataMap)) {
            return result;
        }
        if (!replaceImgData(strDataMap)) {
            return result;
        }
        result = true;
        return result;
    }

    /**
     * 使用数据库中的数据动态填充 word 表格
     * 表格中的数据格式是 List，直接进行填充。
     * 表格数据格式 List &lt; List &lt; String>>
     *
     * @param tblDataMap
     * @return
     */
    private boolean replaceTableSimple(Map<String, List> tblDataMap) {
        boolean result = false;
        NodeList tbNodeList = doc.getElementsByTagName("w:tbl");
        for (Map.Entry<String, List> tbEntry : tblDataMap.entrySet()) {
            if (tbEntry.getValue() == null) {
                System.out.println(tbEntry.getKey() + "Empty table");
                return false;
            }
            String tblKey = tbEntry.getKey();
            ArrayList<List> tblData = (ArrayList<List>) tbEntry.getValue();
            //表格序号
            int tbIndex = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("@") + 1, tblKey.lastIndexOf("$"))) - 1;
            //从第几行开始填充  @1$2：表示对第 1 个表格进行替换，从第 2 行开始
            int fromRow = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("$") + 1)) - 1;
            /** 输入数据进行验证，并进行填充 */
            if (tbIndex >= 0 && tbIndex < tbNodeList.getLength() && (fromRow == 0 || fromRow == 1)) {
                Element tbElement = (Element) tbNodeList.item(tbIndex);
                NodeList trNodeList = tbElement.getElementsByTagName("w:tr");
                if (fromRow == 1 && trNodeList.getLength() < 2) {
                    return false;
                }
                /** 向需要填充数据的表格添加数据行 */
                for (int i = 0; i < tblData.size() - 1; i++) {
                    oneRowCreateBeforePlaceHolder(tbElement, fromRow);
                }
                for (int i = 0; i < tblData.size(); i++) {
                    if (tblData.get(i) == null || tblData.get(i).size() == 0) {
                        continue;
                    }
                    ArrayList<String> rowData = (ArrayList<String>) tblData.get(i);
                    Element rowNode = (Element) trNodeList.item(i + fromRow);
                    NodeList tcNodeList = rowNode.getElementsByTagName("w:tc");
                    for (int j = 0; j < rowData.size() && j < tcNodeList.getLength(); j++) {
                        Element tcNode = (Element) tcNodeList.item(j);
                        String tcData = rowData.get(j);
                        cellDataFill(tcNode, tcData);
                    }
                }
            } else {
                return false;
            }
        }
        return result;
    }

    /**
     * 使用数据库中的数据动态填充 word 表格
     * 表格数据格式 List<Map<String, String>>
     *
     * @param tblDataMap
     * @return
     */

    private boolean replaceTableMap(Map<String, List> tblDataMap) {
        NodeList tbNodeList = doc.getElementsByTagName("w:tbl");
        for (Map.Entry<String, List> tbEntry : tblDataMap.entrySet()) {
            if (tbEntry.getValue() == null) {
                System.out.println(tbEntry.getKey() + "Empty table");
                return false;
            }
            String tblKey = tbEntry.getKey();
            List<Map> tblData = tbEntry.getValue();
            //表格序号
            int tbIndex = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("@") + 1, tblKey.lastIndexOf("$"))) - 1;
            //从第几行开始填充  @1$2：表示对第 1 个表格进行替换，从第 2 行开始
            int fromRow = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("$") + 1)) - 1;
            /** 输入数据进行验证，并进行填充 */
            if (tbIndex >= 0 && tbIndex < tbNodeList.getLength() && (fromRow == 0 || fromRow == 1)) {
                Element tbElement = (Element) tbNodeList.item(tbIndex);
                NodeList trNodeList = tbElement.getElementsByTagName("w:tr");
                if (fromRow >= trNodeList.getLength()) {
                    return false;
                }
                /** 向需要填充数据的表格添加数据行 */
                for (int i = 0; i < tblData.size() - 1; i++) {
                    oneRowCreateBeforePlaceHolder(tbElement, fromRow);
                }
                /** 对表格进行填充数据 */
                for (int i = 0; i < tblData.size(); i++) {
                    if (tblData.get(i) == null || tblData.get(i).size() == 0) {
                        continue;
                    }
                    Map<String, String> rowData = tblData.get(i);
                    Element rowNode = (Element) trNodeList.item(i + fromRow);
                    NodeList tcNodeList = rowNode.getElementsByTagName("w:tc");
                    for (int j = 0; j < rowData.size() && j < tcNodeList.getLength(); j++) {
                        Element tcEmelent = (Element) tcNodeList.item(j);
                        replaceWTContext(tcEmelent, rowData);
                    }
                }
            } else {
                return false;
            }

        }
        return true;
    }


    private void cellDataFill(Element tcNode, String tcData) {
        NodeList wtNodeList = tcNode.getElementsByTagName("w:t");
        if (wtNodeList.getLength() > 0) {
            Element wtNode = (Element) wtNodeList.item(0);
            if (wtNode.hasChildNodes()) {
                wtNode.getFirstChild().setNodeValue(tcData);
            }
        }
    }

    /**
     * 将tbl 中的key 对应的表获取要填充的占位符，将表中第一行或第二行的 占位符填充到 List 中。
     *
     * @param tbl
     */
    public boolean getTableKeyList(Map<String, List<String>> tbl) {
        NodeList tblNodeList = doc.getElementsByTagName("w:tbl");
        for (Map.Entry<String, List<String>> entry : tbl.entrySet()) {
            String tblKey = entry.getKey();
            //表格序号
            int tbIndex = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("@") + 1, tblKey.lastIndexOf("$")));
            //从第几行开始填充  @1$2：表示对第 1 个表格进行替换，从第 2 行开始
            int fromRow = Integer.parseInt(tblKey.substring(tblKey.lastIndexOf("$") + 1)) - 1;
            if (!(fromRow == 0 || fromRow == 1)) {
                return false;
            }
            if (tbIndex - 1 >= 0 && tbIndex - 1 < tblNodeList.getLength()) {
                Element tbNode = (Element) tblNodeList.item(tbIndex - 1);
                NodeList trNodeList = tbNode.getElementsByTagName("w:tr");
                Element trNode = null;
                if (!(trNodeList.getLength() >= fromRow)) {
                    return false;
                }
                trNode = (Element) trNodeList.item(fromRow);
                List<String> trPlaceHolders = getTrCellName(trNode);
                if (trPlaceHolders.isEmpty()) {
                    return false;
                }
                entry.setValue(trPlaceHolders);
            } else {
                return false;
            }
        }
        return true;
    }

    private List<String> getTrCellName(Element trNode) {
        List<String> result = new ArrayList<String>();
        NodeList tcNodeList = trNode.getElementsByTagName("w:tc");
        for (int i = 0; i < tcNodeList.getLength(); i++) {
            Element tcNode = (Element) tcNodeList.item(i);
            NodeList wtNodeList = tcNode.getElementsByTagName("w:t");
            for (int j = 0; j < wtNodeList.getLength(); j++) {
                Element wtNode = (Element) wtNodeList.item(j);
                if (wtNode.hasChildNodes() && matched(wtNode.getFirstChild().getNodeValue())) {
                    result.add(getPlaceStr(wtNode.getFirstChild().getNodeValue()));
                    break;
                }
            }
        }
        return result;
    }

    /**
     * 根据数据库的数据在xml模板中动态添加行
     *
     * @param tblElement 指定的表格
     */
    private boolean oneRowCreate(Element tblElement) {
        if (tblElement != null) {
            NodeList trNodeList = tblElement.getElementsByTagName("w:tr");
            Element trElement = null;
            if (trNodeList.getLength() == 1) {
                trElement = (Element) trNodeList.item(0);
            } else if (trNodeList.getLength() > 1) {
                trElement = (Element) trNodeList.item(1);
            } else {
                System.out.println("Talbe Not Exist!");
                return false;
            }
            //克隆w:tr节点及其子节点
            Element newTrElement = (Element) trElement.cloneNode(true);
            tblElement.appendChild(newTrElement);
            return true;
        }
        return false;
    }

    /**
     * 根据数据库的数据在xml模板中动态添加行，在表格占位符之前添加
     *
     * @param tblElement 指定的表格
     */
    private boolean oneRowCreateBeforePlaceHolder(Element tblElement, int fromRow) {
        if (tblElement != null) {
            NodeList trNodeList = tblElement.getElementsByTagName("w:tr");
            Element trElement = null;
            trElement = (Element) trNodeList.item(fromRow);
            //克隆w:tr节点及其子节点
            Element newTrElement = (Element) trElement.cloneNode(true);
            tblElement.insertBefore(newTrElement, trElement);
            return true;
        }
        return false;
    }


    public boolean doc2XmlFile(String filename) {
        boolean flag = true;
        try {
            TransformerFactory transFactory = TransformerFactory.newInstance();
            Transformer transformer = transFactory.newTransformer();
            DOMSource source = new DOMSource();
            source.setNode(doc);
            StreamResult result = new StreamResult();
            FileOutputStream fileOutputStream = new FileOutputStream(filename);
            result.setOutputStream(fileOutputStream);
            transformer.transform(source, result);
            fileOutputStream.close();
        } catch (Exception ex) {
            flag = false;
            ex.printStackTrace();
        }
        return flag;
    }

    /**
     * 获取文档的 body 字符串
     *
     * @return
     */
    public XmlBody getXmlBody() {
        String bodyContent = null;
        String fileContentStr = com.sse.demo2.service.util.FileUtilHelper.getFileContentStr(filename);
        String end = "</w:body>";
        int bodyStart = fileContentStr.indexOf("<w:body>");
        int bodyEnd = fileContentStr.indexOf(end);
        if (bodyStart != -1 && bodyEnd != -1) {
            bodyContent = fileContentStr.substring(bodyStart, bodyEnd + end.length());
        }
        Matcher matcher = Pattern.compile("<w:document(.*?)>").matcher(fileContentStr);
        XmlBody body = null;
        if (matcher.find()) {
            String namespace = matcher.group(1);
            if (bodyContent != null) {
                body = new XmlBody();
                body.setContent(bodyContent);
                body.setNamespace(namespace);
            }
        }
        return body;
    }

    /**
     * 对 xml 文件进行解析获取所有图片数据
     *
     * @return
     */
    public List<XmlPicture> getXmlAllPictures() {
        List<XmlPicture> allPictures = new ArrayList<>();
        NodeList relationshipsNodeList = doc.getElementsByTagName("Relationships");
        if (!(relationshipsNodeList.getLength() >= 2)) {
            return allPictures;
        }
        Element relationShip = (Element) relationshipsNodeList.item(1);
        NodeList docRelationNodeList = relationShip.getElementsByTagName("Relationship");
        Map<String, XmlPicture> oriPictureMap = new HashMap<>();
        for (int i = 0; i < docRelationNodeList.getLength(); i++) {
            Element relation = (Element) docRelationNodeList.item(i);
            String fileTarget = null;
            if (relation.hasAttribute("Target")
                    && XmlPictureType.isPictureType((fileTarget = relation.getAttribute("Target")))) {
                if (!relation.hasAttribute("Id")) {
                    continue;
                }
                String id = relation.getAttribute("Id");
                int type = XmlPictureType.getPictureType(fileTarget);
                XmlPicture pic = new XmlPicture();
                pic.setType(type);
                pic.setRelationId(id);
                pic.setTarget(fileTarget);
                oriPictureMap.put("/word/" + fileTarget, pic);
            }
        }
        cleanPicData(oriPictureMap);
        for (Map.Entry<String, XmlPicture> entry : oriPictureMap.entrySet()) {
            allPictures.add(entry.getValue());
        }
        return allPictures;
    }

    /**
     * 对 oriPicMap 中的数据填充对应图片的 byte[]；清除没有图片数据的映射关系
     *
     * @param oriPicMap
     */
    private void cleanPicData(Map<String, XmlPicture> oriPicMap) {
        NodeList pkgPartNodeList = doc.getElementsByTagName("pkg:part");
        for (int i = 0; i < pkgPartNodeList.getLength(); i++) {
            Element pkgNode = (Element) pkgPartNodeList.item(i);
            String imgName;
            if (pkgNode.hasAttribute("pkg:name")
                    && oriPicMap.containsKey((imgName = pkgNode.getAttribute("pkg:name")))) {
                NodeList binDataNodeList = pkgNode.getElementsByTagName("pkg:binaryData");
                if (binDataNodeList.getLength() > 0) {
                    Element binNode = (Element) binDataNodeList.item(0);
                    if (binNode.hasChildNodes()) {
                        String imgBas64 = binNode.getFirstChild().getNodeValue();
                        imgBas64 = imgBas64.replace("\n", "");
                        oriPicMap.get(imgName).setData(ImgUtil.getBytes(imgBas64));
                    }
                }
            }
        }
        Iterator<Map.Entry<String, XmlPicture>> it = oriPicMap.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry<String, XmlPicture> next = it.next();
            if (next.getValue().getData() == null) {
                it.remove();
            }
        }
    }

    private static void testMyReplaceDoc() throws ParserConfigurationException, SAXException, IOException {
        Map<String, Object> mark1 = new HashMap<>();
        mark1.put("ONE", "YI");
        mark1.put("TWO", "ER");
        mark1.put("THREE", "SAN");
        mark1.put("FOUR", "SI");
        mark1.put("character", "特点map");
        mark1.put("article2", "第二篇map");
        mark1.put("end", "结束了map");
        List<String> row1 = new ArrayList<>();
        row1.add("1");
        row1.add("2");
        row1.add("3");
        List<String> row2 = new ArrayList<>();
        row2.add("4");
        row2.add("5");
        List<List> tbData = new ArrayList<>();
        tbData.add(row1);
        tbData.add(row2);
        mark1.put("@1$2", tbData);

        String inFilepath = XmlDomHandler.class.getClassLoader().getResource("").getPath() + "doc/temp7.xml";
        String outFilepath = XmlDomHandler.class.getClassLoader().getResource("").getPath() + "doc/out.xml";
        XmlDomHandler handler = new XmlDomHandler(inFilepath);

        handler.renderSimple(mark1);
        handler.doc2XmlFile(outFilepath);
    }

    private static void testMyReplaceDocMap() throws ParserConfigurationException, SAXException, IOException {
        Map<String, Object> mark1 = new HashMap<>();
        mark1.put("hi", "HI");
        mark1.put("name", "sse");
        mark1.put("address", "shanghai");
        mark1.put("character", "特点map");
        mark1.put("article2", "第二篇map");
        mark1.put("end", "结束了map");
        mark1.put("lineone", "第一行");
        mark1.put("linetwo", "第二行");
        mark1.put("linethree", "第三行");

        Map<String, Object> row1 = new HashMap<>();
        row1.put("one", "YI");
        row1.put("two", "ER");
        row1.put("three", "SAN");
        Map<String, Object> row2 = new HashMap<>();
        row2.put("one", "YI2");
        row2.put("two", "ER2");
        row2.put("three", "SAN2");
        List<Map> tbData = new ArrayList<>();
        tbData.add(row1);
        tbData.add(row2);
//        mark1.put("@1$2", tbData);
        String inFilepath = XmlDomHandler.class.getClassLoader().getResource("").getPath() + "doc/temp1.xml";
        String outFilepath = XmlDomHandler.class.getClassLoader().getResource("").getPath() + "doc/temp1.doc";
        XmlDomHandler handler = new XmlDomHandler(inFilepath);
        handler.renderMap(mark1);
        handler.flushFile(outFilepath);
    }

    private static void testExtractPicture() throws ParserConfigurationException, SAXException, IOException {
        String inFilepath = XmlDomHandler.class.getClassLoader().getResource("").getPath() + "doc/temp1.xml";
        XmlDomHandler handler = new XmlDomHandler(inFilepath);
        XmlBody xmlBody = handler.getXmlBody();
        List<XmlPicture> allPictures = handler.getXmlAllPictures();
    }

    public static void main(String[] args) throws IOException, SAXException, ParserConfigurationException {

        testMyReplaceDocMap();
    }

}
