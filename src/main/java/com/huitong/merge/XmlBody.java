package com.huitong.merge;

/**
 * @author pczhao
 * @email
 * @date 2018-10-18 19:16
 */

public class XmlBody {
    /** xml 文件的 body 内容 */
    private String content;
    /** xml 文件的 body namespace */
    private String namespace;

    public XmlBody() {
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getNamespace() {
        return namespace;
    }

    public void setNamespace(String namespace) {
        this.namespace = namespace;
    }

    /**
     * 将 body 内容添加 namespace 以便 XWPFDocument 解析时验证
     * @return
     */
    public String getCTBodyStr() {
        if (content == null || namespace == null) {
            return null;
        }
        int bdStartlast = content.indexOf(">");
        StringBuilder sb = new StringBuilder();
        sb.append("<w:body ");
        sb.append(namespace);
        sb.append(" >");
        sb.append(content.substring(bdStartlast + 1));
        return sb.toString();
    }
}
