package com.huitong.wordreplace;

import java.util.HashMap;
import java.util.Map;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-10-28 7:58
 */

public class Demo1 {

    public static void main(String[] args) {

        Map<Long, Long> map = new HashMap<>();
        long i = 0;
        try {
            for (i = 0; i<Long.MAX_VALUE; i++) {
                map.put(i,i);
            }
        } catch (Exception e){
            e.printStackTrace();
            System.out.println(i);
        }

        System.out.println(map.size());
    }
}
