package com.itwobyte.framework.fileHandler.imp;

import cn.hutool.crypto.digest.MD5;
import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.itwobyte.common.constant.CacheConstants;
import com.itwobyte.common.utils.StringUtils;
import com.itwobyte.common.utils.file.FileUtils;
import com.itwobyte.framework.fileHandler.FileFactory;
import org.redisson.api.RMapCache;
import org.redisson.api.RedissonClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;

@Service("docx")
public class Docx2PdfService implements FileFactory {

    private static final Logger log = LoggerFactory.getLogger(Docx2PdfService.class);


    @Autowired
    private RedissonClient redissonClient;

    @Override
    public void filePreview(String source,String md5,OutputStream outputStream, HttpServletResponse response) throws Exception {
        // 验证License
        if (!isWordLicense()) {
            return ;
        }

        String downloadName = StringUtils.substringAfterLast(source, "/");
        response.setContentType("application/x-msdownload;charset=utf-8");
        response.setHeader("Content-Disposition",
            "attachment; filename=\"" + URLEncoder.encode(downloadName.substring(0,downloadName.lastIndexOf("."))+".pdf","UTF-8")+ "\"");

        // 字体
        String linux = System.getProperty("os.name");
        if (StringUtils.isNotEmpty(linux) && linux.contains("Linux")) {
            String fontPath = "/usr/share/fonts";
            if (!new File(fontPath).exists()) {
                log.debug("linux系统,字体目录不存在" + fontPath);
            }
            FontSettings.setFontsFolder(fontPath, true);
        }


        String target  =  source.substring(0,source.lastIndexOf("."))+ "_preview.pdf";

        FileOutputStream os = null;
        try {
            // Address是将要被转化的word文档
            Document doc = new Document(Files.newInputStream(Paths.get(source)));

            if(StringUtils.isNotEmpty(md5)){

                File file = new File(target);
                os = new FileOutputStream(file);
                doc.save(os, com.aspose.words.SaveFormat.PDF);

                RMapCache<String, String> mapCache = redissonClient.getMapCache(CacheConstants.FILE_PREVIEW_KEY);
                mapCache.fastPut(md5,target);
                FileUtils.writeBytes(target,response.getOutputStream());
                return;
            }

            // 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
            doc.save(outputStream, com.aspose.words.SaveFormat.PDF);
        } catch (Exception e) {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            log.debug(e.toString());
            throw e;
        }


    }

    @Override
    public void filePreviewCache(String source,String md5,OutputStream outputStream, HttpServletResponse response, HttpServletRequest request) throws Exception {

        RMapCache<String, String> mapCache = redissonClient.getMapCache(CacheConstants.FILE_PREVIEW_KEY);
        if(mapCache!=null && StringUtils.isNotEmpty(md5) &&  mapCache.get(md5)!=null){
            String path = mapCache.get(md5);
            System.out.println(path);
            File file =new File(path);
            if(file.exists()){
                String downloadName = StringUtils.substringAfterLast(path, "/");
                response.setContentType("application/x-msdownload;charset=utf-8");
                response.setHeader("Content-Disposition",
                    "attachment; filename=\"" + URLEncoder.encode(downloadName.substring(0,downloadName.lastIndexOf("."))+".pdf","UTF-8")+ "\"");
                FileUtils.writeBytes(path, response.getOutputStream());
            }else {
                this.filePreview(source,md5,outputStream,response);
            }
        }else {
            this.filePreview(source,md5,outputStream,response);
        }
    }

    @Override
    public String file2PdfAndSave(String source, String target, String md5) throws Exception {
        // 验证License
        if (!isWordLicense()) {
            return null;
        }
        target = getPreview(target);

        String linux = System.getProperty("os.name");
        if (StringUtils.isNotEmpty(linux) && linux.contains("Linux")) {
            String fontPath = "/usr/share/fonts";
            if (!new File(fontPath).exists()) {
                log.debug("linux系统,字体目录不存在" + fontPath);
            }
            FontSettings.setFontsFolder(fontPath, true);
        }
        FileOutputStream os = null;
        try {
            String path = target.substring(0, target.lastIndexOf('/'));
            File file = new File(path);
            // 创建文件夹
            if (!file.exists()) {
                file.mkdirs();
            }
            // 新建一个空白pdf文档
            file = new File(target);
            os = new FileOutputStream(file);
            // Address是将要被转化的word文档
            Document doc = new Document(Files.newInputStream(Paths.get(source)));
            // 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
            doc.save(os, com.aspose.words.SaveFormat.PDF);
            os.close();
        } catch (Exception e) {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            log.debug(e.toString());
            throw e;
        }
        return target;
    }

    /**
     * 验证 Aspose.word 组件是否授权
     * 无授权的文件有水印和试用标记
     */
    public static boolean isWordLicense() {
        boolean result = false;
        try {
            // 避免文件遗漏
            String licensexml = "<License>\n" +
                "<Data>\n" +
                "<Products>\n" +
                "<Product>Aspose.Total for Java</Product>\n" +
                "</Products>\n" +
                "<EditionType>Enterprise</EditionType>\n" +
                "<SubscriptionExpiry>20991231</SubscriptionExpiry>\n" +
                "<LicenseExpiry>20991231</LicenseExpiry>\n" +
                "<SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n" +
                "</Data>\n" +
                "<Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature>\n" +
                "</License>";
            InputStream inputStream = new ByteArrayInputStream(licensexml.getBytes());
            com.aspose.words.License license = new com.aspose.words.License();
            license.setLicense(inputStream);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static String getPreview(String target){
        String s = target.substring(0, target.lastIndexOf(".")) + "_preview.pdf";

        return  s;
    }
}
