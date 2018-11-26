package com.huitong.merge;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.*;

/**
 * @author ZHAOPENGCHENG
 * @email
 * @date 2018-11-22 22:00
 */

public class MultiThreadMerge {

    public boolean merge(final String destFilename, List<String> mergeDocs) {
        int threadCount = mergeDocs.size() / 90 + 1;
        String[][] threadDocs = new String[threadCount][mergeDocs.size() / threadCount + 1];
        List<String> tmergeDocs = new LinkedList<>();
        tmergeDocs.addAll(mergeDocs);

        HashSet<String> globalAlreadyFiles = new HashSet<>(128);
        for (int i = 0; i < threadCount; i++) {
            HashSet<String> threadAlreadyFiles = new HashSet<>(128);
            int everyThreadDocsCount = mergeDocs.size() / threadCount + 1;
            for (int j = 0; j < everyThreadDocsCount && !tmergeDocs.isEmpty(); j++) {
                threadDocs[i][j] = tmergeDocs.remove(0);
                if (globalAlreadyFiles.contains(threadDocs[i][j])) {
                    String threadSafeFilename = getThreadSafeFilename(threadDocs[i][j], i);
                    if (!threadAlreadyFiles.contains(threadSafeFilename)) {
                        if (!FileUtil.copyFile(threadSafeFilename, threadDocs[i][j])) {
                            return false;
                        }
                    }
                    threadDocs[i][j] = threadSafeFilename;
                }
                threadAlreadyFiles.add(threadDocs[i][j]);
            }
            globalAlreadyFiles.addAll(threadAlreadyFiles);
        }
        /** 各个线程的合并生成文件 */
        List<String> genFilenames = new ArrayList<>();
        /** 各个县城的返回结果 */
        List<Future<Boolean>> threadResults = new ArrayList<>();

        ExecutorService executorService = Executors.newFixedThreadPool(threadCount);
        for (int i = 0; i < threadCount; i++) {
            ArrayList<String> singThreadDocs = new ArrayList<>();
            for (String s : threadDocs[i]) {
                if (!StringUtils.isBlank(s)) {
                    singThreadDocs.add(s);
                }
            }
            if (!singThreadDocs.isEmpty()) {
                genFilenames.add(getThreadSafeFilename(destFilename, i));
                threadResults.add(executorService.submit(new MergeDocThread(singThreadDocs, genFilenames.get(i))));
            }
        }
        for (Future<Boolean> result : threadResults) {
            try {
                if (!result.get()) {
                    return false;
                }
            } catch (InterruptedException e) {
                e.printStackTrace();
                Thread.currentThread().interrupt();
                result.cancel(true);
                executorService.shutdownNow();
                return false;
            } catch (ExecutionException e) {
                executorService.shutdownNow();
                e.getCause().printStackTrace();
                return false;
            }
        }
        executorService.shutdown();
        MergeDocHandler handler = new MergeDocHandler();
        boolean mergeResult = false;
        try {
            mergeResult = handler.iMergeMixedFileList2Docx(destFilename, genFilenames);
        } catch (InvalidFormatException | XmlException | IOException | ParserConfigurationException | SAXException e) {
            e.printStackTrace();
            return false;
        }
        return mergeResult;
    }

    private String getThreadSafeFilename(final String oriFilename, final int threadNum) {
        StringBuilder sb = new StringBuilder();
        int sepIndex = oriFilename.lastIndexOf(".");
        sb.append(oriFilename.substring(0, sepIndex));
        sb.append(threadNum);
        sb.append(oriFilename.substring(sepIndex));
        return sb.toString();
    }
}
