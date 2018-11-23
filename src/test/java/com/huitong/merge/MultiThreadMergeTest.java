package com.huitong.merge;

import org.junit.Assert;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-11-23 0:03
 */

public class MultiThreadMergeTest {

    @Test
    public void testMultiThreadMerge() {
        String inFile1 = Class.class.getClass().getResource("/").getPath() + "doc/dom1.doc";
        String outfile = Class.class.getClass().getResource("/").getPath() + "doc/0out.docx";
        List<String> docs = new ArrayList<>();
        for (int i = 0; i < 500; i++) {
            docs.add(inFile1);
        }
        MultiThreadMerge handler = new MultiThreadMerge();
        boolean result = handler.merge(outfile, docs);
        Assert.assertTrue(result);
    }

}
