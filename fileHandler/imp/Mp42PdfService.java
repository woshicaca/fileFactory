package com.itwobyte.framework.fileHandler.imp.fileHandler.imp;

import com.itwobyte.common.utils.StringUtils;
import com.itwobyte.framework.config.MedioHttpRequestHandler;
import org.redisson.api.RedissonClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Service("mp4")
public class Mp42PdfService implements FileFactory {

    private static final Logger log = LoggerFactory.getLogger(Mp42PdfService.class);


    @Autowired
    private RedissonClient redissonClient;

    @Resource
    private MedioHttpRequestHandler medioHttpRequestHandler;

    @Override
    public void filePreview(String source,String md5,OutputStream outputStream, HttpServletResponse response) throws Exception {
    }

    @Override
    public void filePreviewCache(String source, String md5, OutputStream outputStream, HttpServletResponse response, HttpServletRequest request) throws Exception {
        Path path = Paths.get(source);
        if (Files.exists(path)) {
            String mimeType = Files.probeContentType(path);
            if (!StringUtils.isEmpty(mimeType)) {
                response.setContentType(mimeType);
            }
            request.setAttribute(MedioHttpRequestHandler.ATTR_FILE, path);
            try {
                medioHttpRequestHandler.handleRequest(request, response);
            }catch (Exception e){

            }

        } else {
            response.setStatus(HttpServletResponse.SC_NOT_FOUND);
            response.setCharacterEncoding(StandardCharsets.UTF_8.toString());
        }
    }

    @Override
    public String file2PdfAndSave(String source, String target, String md5) throws Exception {
      return "";
    }

}
