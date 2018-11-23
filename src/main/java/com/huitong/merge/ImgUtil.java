package com.huitong.merge;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.lang3.StringUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author pczhao
 * @email
 * @date 2018-10-11 13:59
 */

public class ImgUtil {
    /**
     * 将图片转换成Base64编码
     * @param imgFilename
     * @return
     */
    public static String getImgStr(final String imgFilename) {
        //将图片文件转化为字节数组字符串，并对其进行Base64编码处理
        String result = null;
        byte[] data = getImgBytes(imgFilename);
        if (data != null) {
            result = Base64.encodeBase64String(data);
        }
        return result;
    }

    public static byte[] getImgBytes(final String imgFilename) {
        FileInputStream inputStream = null;
        byte[] data = null;
        try {
            inputStream = new FileInputStream(imgFilename);
            data = new byte[inputStream.available()];
            inputStream.read(data);
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    public static String getImgStr(byte[] imgData) {
    	String result = null;
    	if(imgData != null) {
    		result = Base64.encodeBase64String(imgData);
    	}
    	return result;
    }

    /**
     * 对字节数组字符串进行Base64解码并生成图片
     * @param imgStr 图片数据
     * @param desImgFilePath 保存图片全路径地址
     * @return
     */
    public static boolean genImg(final String imgStr, final String desImgFilePath) {
        boolean result = false;
        if (StringUtils.isEmpty(imgStr)) {
            return result;
        }
        byte[] data = Base64.decodeBase64(imgStr);
        for (int i=0; i < data.length; i++) {
            if (data[i] < 0) {
                data[i] += 256;
            }
        }
        //生成图片
        try {
            FileOutputStream os = new FileOutputStream(desImgFilePath);
            os.write(data);
            os.close();
            result = true;
        } catch (IOException e) {
            e.printStackTrace();
            result = false;
        }
        return result;
    }

    public static byte[] getBytes(String imgBas64) {
        if (imgBas64 == null) {
            return null;
        }
        return Base64.decodeBase64(imgBas64);
    }


}
