package com.itwobyte.framework.fileHandler.imp.fileHandler;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;

public interface FileFactory {

    public void filePreview(String source,String md5, OutputStream outputStream, HttpServletResponse response) throws Exception;


    public void filePreviewCache(String source, String md5, OutputStream outputStream, HttpServletResponse response, HttpServletRequest request) throws Exception;



    String file2PdfAndSave(String source, String target, String md5) throws Exception;
}
