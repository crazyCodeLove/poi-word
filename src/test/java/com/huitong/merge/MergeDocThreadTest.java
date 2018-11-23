package com.huitong.merge;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-11-23 0:26
 */

public class MergeDocThreadTest {

    @Test
    public void testMergeDoc() {
        String inFile1 = Class.class.getClass().getResource("/").getPath() + "doc/1.docx";
        String outfile = Class.class.getClass().getResource("/").getPath() + "doc/0out.docx";
        List<String> docs = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            docs.add(inFile1);
        }
        ExecutorService executor = Executors.newCachedThreadPool();
        Future<Boolean> result = executor.submit(new MergeDocThread(docs, outfile));
        executor.shutdown();
        try {
            System.out.println(result.get());
        } catch (InterruptedException|ExecutionException e) {
            e.printStackTrace();
        }
    }
}
