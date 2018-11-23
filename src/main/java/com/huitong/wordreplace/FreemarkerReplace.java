package com.huitong.wordreplace;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-10-13 7:43
 */

public class FreemarkerReplace {

    public static String replaceKeyStart(String ori) {
        StringBuilder sb = new StringBuilder();

        if (ori.contains("${")) {
            int start = -1;
            int nextStart = -1, nextEnd = -1, lastStart = -1;
            while (start + 1 < ori.length() && (nextStart = ori.indexOf("${", start +1)) != -1) {
                sb.append(ori.substring(start + 1, nextStart));
                nextEnd = ori.indexOf("}", nextStart + 1);
                lastStart = ori.indexOf("${", nextStart + 1);
                if (nextEnd != -1) {
                    /** 之后的字符串中有 ${ 和 } */
                    if (lastStart == -1) {
                        /** ${ 后的子串中仅有一个 } 没有 ${了*/
                        sb.append(ori.substring(start + 1));
                        start = nextEnd;
                    } else {
                        /** ${ 后的子串中 既有 ${ 也有 } */
                        if (nextEnd < lastStart) {
                            /** ${后的子串中 } 在 ${ 之前  */
                            sb.append(ori.substring(start+1, lastStart));
                            start = lastStart -1;
                        } else {
                            sb.append(ori.substring(start +1, nextStart));
                            sb.append("${r'${'}");
                            sb.append(ori.substring(nextStart + 2, nextEnd + 1));
                            start = nextEnd;
                        }
                    }
                } else {
                    /** 没有找到与 ${ 配对的 } */
                    sb.append(ori.substring(start + 1, nextStart));
                    sb.append("${r'${'}");
                    if (nextStart + 2 < ori.length()) {
                        sb.append(ori.substring(nextStart + 2));
                    }
                    start = ori.length() - 1;
                }
            }
        } else {
            sb.append(ori);
        }
        return sb.toString();
    }

    public static void main(String[] args) {
        String ori = "${${${}}";
        System.out.println(replaceKeyStart(ori));
    }
}
