package com.huitong.merge;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.Document;

import java.util.HashMap;
import java.util.Map;

/**
 * @author pczhao
 * @email
 * @date 2018-10-18 13:36
 */

public final class XmlPictureType {
    public static Map<String, Integer> picTypeMap = new HashMap<>();
    static {
        picTypeMap.put(".jpg", Document.PICTURE_TYPE_JPEG);
        picTypeMap.put(".jpeg", Document.PICTURE_TYPE_JPEG);
        picTypeMap.put(".jpe", Document.PICTURE_TYPE_JPEG);
        picTypeMap.put(".jfif", Document.PICTURE_TYPE_JPEG);
        picTypeMap.put(".gif", Document.PICTURE_TYPE_GIF);
        picTypeMap.put(".tif", Document.PICTURE_TYPE_TIFF);
        picTypeMap.put(".tiff", Document.PICTURE_TYPE_TIFF);
        picTypeMap.put(".png", Document.PICTURE_TYPE_PNG);
        picTypeMap.put(".bpm", Document.PICTURE_TYPE_BMP);
        picTypeMap.put(".dib", Document.PICTURE_TYPE_DIB);
    }

    /**
     * 判断文件名是否是图片类型
     * @param filename
     * @return
     */
    public static boolean isPictureType(String filename) {
        boolean result = false;
        if (StringUtils.isBlank(filename) || !filename.contains(".")) {
            return result;
        }
        filename = filename.trim().toLowerCase();
        int start = filename.lastIndexOf(".");
        return picTypeMap.containsKey(filename.substring(start));
    }

    public static int getPictureType(String filename) {
        if (!isPictureType(filename)) {
            throw new RuntimeException("不是图片类型，不能获得对应的类型");
        }
        filename = filename.trim().toLowerCase();
        int start = filename.lastIndexOf(".");
        return picTypeMap.get(filename.substring(start));
    }

    public static void main(String[] args) {
        String filename = "media/image2..JPeg";
        System.out.println(isPictureType(filename));
    }
}
