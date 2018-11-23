package com.huitong.merge;

import java.util.List;
import java.util.concurrent.Callable;

/**
 * @author pczhao
 * @email pczhao@sse.com.cn
 * @date 2018-11-22 20:15
 */

public class MergeDocThread implements Callable<Boolean> {
    private List<String> mergeList;
    private String destFilename;

    public MergeDocThread(List<String> mergeList, String destFilename) {
        this.mergeList = mergeList;
        this.destFilename = destFilename;
    }

    @Override
    public Boolean call() throws Exception {
        MergeDocHandler handler = new MergeDocHandler();
        boolean result = handler.iMergeMixedFileList2Docx(destFilename, mergeList);
        return result;
    }
}
