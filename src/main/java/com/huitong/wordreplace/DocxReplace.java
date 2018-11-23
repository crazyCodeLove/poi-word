package com.huitong.wordreplace;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-10-10 7:13
 */

public class DocxReplace {

    public static void main(String[] args) {
        try {
            transfer();
        } catch (IOException e) {
            e.printStackTrace();
        }
        String now = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
        System.out.println("完成时间：" + now);

    }

    public static void transfer() throws IOException {
        String inFilename = DocxReplace.class.getClassLoader().getResource("").getPath() + "doc/tldemo1.docx";
        String outFilename = DocxReplace.class.getClassLoader().getResource("").getPath() + "doc/out.docx";

        XWPFTemplate template = XWPFTemplate.compile(inFilename).render(getData());
        FileOutputStream outputStream = new FileOutputStream(outFilename);
        template.write(outputStream);

        outputStream.flush();
        outputStream.close();
        template.close();
    }

    public static Map<String, String> getData() {
        HashMap<String, String> result = new HashMap<>();
        result.put("title1", "正常标题1");
        result.put("Title2", "首字母大写标题2");

        return result;

    }

}
