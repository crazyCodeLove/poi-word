package com.huitong.merge;

/**
 * @author pczhao
 * @email
 * @date 2018-10-18 13:10
 */


public class XmlPicture {
    private byte[] data;
    private int type;
    private String relationId;

    private String target;


    public XmlPicture() {
    }

    public byte[] getData() {
        return data;
    }

    public void setData(byte[] data) {
        this.data = data;
    }

    public int getType() {
        return type;
    }

    public void setType(int type) {
        this.type = type;
    }

    public String getRelationId() {
        return relationId;
    }

    public void setRelationId(String relationId) {
        this.relationId = relationId;
    }

    public String getTarget() {
        return target;
    }

    public void setTarget(String target) {
        this.target = target;
    }
}
