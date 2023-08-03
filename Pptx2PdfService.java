package com.itwobyte.framework.fileHandler.imp;

import cn.hutool.crypto.digest.MD5;
import com.aspose.slides.Presentation;
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

@Service("pptx")
public class Pptx2PdfService implements FileFactory {

    private static final Logger log = LoggerFactory.getLogger(Ppt2PdfService.class);


    @Autowired
    private RedissonClient redissonClient;

    @Override
    public void filePreview(String source,String md5,OutputStream outputStream, HttpServletResponse response) throws Exception {
        if(StringUtils.isEmpty(source)){
            return;
        }
        // 验证License
        if (!getPptLicense()) {
            return ;
        }

        // 字体
        String linux = System.getProperty("os.name");
        if (StringUtils.isNotEmpty(linux) && linux.contains("Linux")) {
            String fontPath = "/usr/share/fonts";
            if (!new File(fontPath).exists()) {
                log.debug("linux系统,字体目录不存在" + fontPath);
            }
            FontSettings.setFontsFolder(fontPath, true);
        }

        String downloadName = StringUtils.substringAfterLast(source, "/");
        response.setContentType("application/x-msdownload;charset=utf-8");
        response.setHeader("Content-Disposition",
            "attachment; filename=\"" + URLEncoder.encode(downloadName.substring(0,downloadName.lastIndexOf("."))+".pdf","UTF-8")+ "\"");



        String target  =  source.substring(0,source.lastIndexOf("."))+ "_preview.pdf";

        FileOutputStream os = null;
        try {
            // Address是将要被转化的word文档
            Presentation pres = new Presentation(Files.newInputStream(Paths.get(source)));

            if(StringUtils.isNotEmpty(md5)){

                File file = new File(target);
                os = new FileOutputStream(file);
                pres.save(os, com.aspose.slides.SaveFormat.Pdf);

                RMapCache<String, String> mapCache = redissonClient.getMapCache(CacheConstants.FILE_PREVIEW_KEY);
                mapCache.fastPut(md5,target);
                FileUtils.writeBytes(target,response.getOutputStream());
                return;
            }
            pres.save(outputStream, com.aspose.slides.SaveFormat.Pdf);
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
        if(mapCache!=null && StringUtils.isNotEmpty(md5) && mapCache.get(md5)!=null){
            String path = mapCache.get(md5);
            File file =new File(path);
            if(file.exists()){
                String downloadName = StringUtils.substringAfterLast(path, "/");
                response.setContentType("application/x-msdownload;charset=utf-8");
                response.setHeader("Content-Disposition",
                    "attachment; filename=\"" + URLEncoder.encode(downloadName.substring(0,downloadName.lastIndexOf("."))+".pdf","UTF-8")+ "\"");
                FileUtils.writeBytes(path, outputStream);
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
        if (!getPptLicense()) {
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
        log.debug("log3");
        FileOutputStream os = null;
        try {
            String path = target.substring(0, target.lastIndexOf("/"));
            File file = new File(path);
            // 创建文件夹
            if (!file.exists()) {
                file.mkdirs();
            }
            // 新建一个空白pdf文档
            file = new File(target);
            os = new FileOutputStream(file);
            // Address是将要被转化的PPT幻灯片
            Presentation pres = new Presentation(Files.newInputStream(Paths.get(source)));
            pres.save(os, com.aspose.slides.SaveFormat.Pdf);
            os.close();
        } catch (Exception e) {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            e.printStackTrace();

            log.debug(e.toString());
            throw e;
        }
        return target;
    }

    private static boolean getPptLicense() {
        boolean result = false;
        try {
            String license =
                "<License>\n" +
                    "  <Data>\n" +
                    "    <Products>\n" +
                    "      <Product>Aspose.Total for Java</Product>\n" +
                    "    </Products>\n" +
                    "    <EditionType>Enterprise</EditionType>\n" +
                    "    <SubscriptionExpiry>20991231</SubscriptionExpiry>\n" +
                    "    <LicenseExpiry>20991231</LicenseExpiry>\n" +
                    "    <SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n" +
                    "  </Data>\n" +
                    "  <Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature>\n" +
                    "</License>";
            InputStream is = new ByteArrayInputStream(license.getBytes("UTF-8"));
            com.aspose.slides.License aposeLic = new com.aspose.slides.License();
            aposeLic.setLicense(is);
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
