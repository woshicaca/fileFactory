package com.itwobyte.framework.fileHandler.imp;

import cn.hutool.crypto.digest.MD5;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.itwobyte.common.constant.CacheConstants;
import com.itwobyte.common.utils.StringUtils;
import com.itwobyte.common.utils.file.FileUtils;
import com.itwobyte.framework.fileHandler.FileFactory;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ReadingOrder;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service("xls")
public class Xsl2XslxService implements FileFactory {

    private static final Logger log = LoggerFactory.getLogger(Xsl2XslxService.class);

    @Autowired
    private RedissonClient redissonClient;



    @Override
    public void filePreview(String source,String md5,OutputStream outputStream, HttpServletResponse response) throws Exception {

        if(StringUtils.isEmpty(source)){
            return;
        }

        //xls2xlsx(source,md5,outputStream,response);

        //创建hssworkbook 操作xls 文件
        POIFSFileSystem fs = new POIFSFileSystem(new File(source));
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fs);

        //创建xssfworkbook 操作xlsx 文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetNum = hssfWorkbook.getNumberOfSheets();


        for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++) {

            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(sheetIndex);

            if (workbook.getSheet(hssfSheet.getSheetName()) == null) {
                XSSFSheet xssfSheet = workbook.createSheet(hssfSheet.getSheetName());
                copySheets(hssfSheet, xssfSheet);
            } else {
                copySheets(hssfSheet, workbook.createSheet(hssfSheet.getSheetName()));
            }


        }
        String downloadName = StringUtils.substringAfterLast(source, "/");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"" +URLEncoder.encode(downloadName,"UTF-8")+"x" + "\"");
        response.setHeader("Accept-Ranges", "bytes");
        response.setHeader("Etag", "W/\"9767057-1323779115364\"");


        String target  =  source.substring(0,source.lastIndexOf("."))+ "_preview.xlsx";

//        xls2xlsx(source,target,outputStream,response);
        FileOutputStream os = null;
        if(StringUtils.isNotEmpty(md5)){
          try{
              File file = new File(target);
              os = new FileOutputStream(file);
              workbook.write(os);

              RMapCache<String, String> mapCache = redissonClient.getMapCache(CacheConstants.FILE_PREVIEW_KEY);
              mapCache.fastPut(md5,target);
              FileUtils.writeBytes(target,response.getOutputStream());

              workbook.close();
              hssfWorkbook.close();
              return;
          }finally {
              if (os != null) {
                  try {
                      os.close();
                  } catch (IOException e1) {
                      e1.printStackTrace();
                  }
              }
          }

        }


        //将复制的xls数据写入到新的xlsx文件中
        workbook.write(outputStream);

        workbook.close();
        hssfWorkbook.close();

    }

    public static String xls2xlsx(String sourse,String md5,OutputStream outputStream, HttpServletResponse response) throws Exception {
        // 验证License
        if (!getXslLicense()) {
            return null;
        }

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
            String target  =  sourse.substring(0,sourse.lastIndexOf("."))+ "_preview.xlsx";
            // Address是将要被转化的excel表格
            Workbook workbook = new Workbook(sourse);
            workbook.save(outputStream, SaveFormat.XLSX);
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
        return sourse;
    }

    private static boolean getXslLicense() {
        boolean result = false;
        try {
            // 凭证
            String license =
                "<License>\n" +
                    "  <Data>\n" +
                    "    <Products>\n" +
                    "      <Product>Aspose.Total for Java</Product>\n" +
                    "      <Product>Aspose.Words for Java</Product>\n" +
                    "    </Products>\n" +
                    "    <EditionType>Enterprise</EditionType>\n" +
                    "    <SubscriptionExpiry>20991231</SubscriptionExpiry>\n" +
                    "    <LicenseExpiry>20991231</LicenseExpiry>\n" +
                    "    <SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n" +
                    "  </Data>\n" +
                    "  <Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature>\n" +
                    "</License>";
            InputStream is = new ByteArrayInputStream(license.getBytes("UTF-8"));
            com.aspose.cells.License asposeLic = new com.aspose.cells.License();
            asposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }



    @Override
    public void filePreviewCache(String source,String md5,OutputStream outputStream, HttpServletResponse response, HttpServletRequest request) throws Exception {
        RMapCache<String, String> mapCache = redissonClient.getMapCache(CacheConstants.FILE_PREVIEW_KEY);
        if(mapCache!=null && StringUtils.isNotEmpty(md5) && mapCache.get(md5)!=null){
            String path = mapCache.get(md5);
            File file =new File(path);
            if(file.exists()){
                String downloadName = StringUtils.substringAfterLast(source, "/");
                response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.setHeader("Content-Disposition", "attachment; filename=\"" +URLEncoder.encode(downloadName,"UTF-8")+"x" + "\"");
                response.setHeader("Accept-Ranges", "bytes");
                response.setHeader("Etag", "W/\"9767057-1323779115364\"");
                FileUtils.writeBytes(path, outputStream);
            }else {
                this.filePreview(source,md5,outputStream,response);
            }
        }else {
            this.filePreview(source,md5,outputStream,response);
        }
    }


    /**
     * xls 文件转换为xlsx文件
     */

    public static String xls2xlsx(String sourceFile) throws IOException {

        //创建hssworkbook 操作xls 文件
        POIFSFileSystem fs = new POIFSFileSystem(new File(sourceFile));
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fs);

        //创建xssfworkbook 操作xlsx 文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        int sheetNum = hssfWorkbook.getNumberOfSheets();

        String xlsxPath = createNewXlsxFilePath(sourceFile);



        for (int sheetIndex = 0; sheetIndex < sheetNum; sheetIndex++) {

            HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(sheetIndex);

            if (workbook.getSheet(hssfSheet.getSheetName()) == null) {
                XSSFSheet xssfSheet = workbook.createSheet(hssfSheet.getSheetName());
                copySheets(hssfSheet, xssfSheet);
            } else {
                copySheets(hssfSheet, workbook.createSheet(hssfSheet.getSheetName()));
            }


            FileOutputStream fileOut = new FileOutputStream(xlsxPath);

            //将复制的xls数据写入到新的xlsx文件中
            workbook.write(fileOut);

            workbook.close();
            hssfWorkbook.close();

        }

        return  xlsxPath;
    }

    //为xlsx创建路径
    public static String createNewXlsxFilePath(String sourceFile){

        StringBuffer fPath = new StringBuffer(sourceFile);
        String flag = "_preview";
        fPath = fPath.insert(fPath.lastIndexOf("."), flag);

        return  fPath.toString()+"x";
    }


    /**
     * 转换为xlsx --创建sheet
     * @param source
     * @param destination
     */

    public static void copySheets(HSSFSheet source, XSSFSheet destination) {

        int maxColumnNum = 0;
        // 获取全部的合并单元格
        List<CellRangeAddress> cellRangeAddressList = source.getMergedRegions();
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            if(i<0){
                continue;
            }
            HSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                // 拷贝行
                copyRow(destination, srcRow, destRow, cellRangeAddressList);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }

        // 拷贝图片
        copyPicture(source, destination);

    }


    /**
     * 转换xlsx --  复制行
     * @param srcRow
     * @param destRow
     */
    public static void copyRow(XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow,
                               List<CellRangeAddress> cellRangeAddressList) {

        // 拷贝行高
        destRow.setHeight(srcRow.getHeight());

        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            if(j<0 ){
                continue;
            }
            HSSFCell oldCell = srcRow.getCell(j);
            XSSFCell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                // 拷贝单元格
                copyCell(oldCell, newCell, destSheet);

                // 获取原先的合并单元格
                CellRangeAddress mergedRegion = getMergedRegion(cellRangeAddressList, srcRow.getRowNum(),
                    (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    // 参照创建合并单元格
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                        mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    destSheet.addMergedRegion(newMergedRegion);
                }
            }
        }

    }



    // 拷贝单元格
    public static void copyCell(HSSFCell oldCell, XSSFCell newCell, XSSFSheet destSheet) {

        HSSFCellStyle sourceCellStyle = oldCell.getCellStyle();
        XSSFCellStyle targetCellStyle = destSheet.getWorkbook().createCellStyle();

        if (sourceCellStyle == null) {
            sourceCellStyle = oldCell.getSheet().getWorkbook().createCellStyle();
        }

        targetCellStyle.setFillForegroundColor(sourceCellStyle.getFillForegroundColor());
        // 设置对齐方式
        targetCellStyle.setAlignment(sourceCellStyle.getAlignment());
        targetCellStyle.setVerticalAlignment(sourceCellStyle.getVerticalAlignment());

        // 设置字体
        XSSFFont xssfFont = destSheet.getWorkbook().createFont();
        HSSFFont hssfFont = sourceCellStyle.getFont(oldCell.getSheet().getWorkbook());
        copyFont(xssfFont, hssfFont);
        targetCellStyle.setFont(xssfFont);
        // 文本换行
        targetCellStyle.setWrapText(sourceCellStyle.getWrapText());

        targetCellStyle.setBorderBottom(sourceCellStyle.getBorderBottom());
        targetCellStyle.setBorderLeft(sourceCellStyle.getBorderLeft());
        targetCellStyle.setBorderRight(sourceCellStyle.getBorderRight());
        targetCellStyle.setBorderTop(sourceCellStyle.getBorderTop());
        targetCellStyle.setBottomBorderColor(sourceCellStyle.getBottomBorderColor());
        targetCellStyle.setDataFormat(sourceCellStyle.getDataFormat());
        targetCellStyle.setFillBackgroundColor(sourceCellStyle.getFillBackgroundColor());
        targetCellStyle.setFillPattern(sourceCellStyle.getFillPattern());

        targetCellStyle.setHidden(sourceCellStyle.getHidden());
        targetCellStyle.setIndention(sourceCellStyle.getIndention());
        targetCellStyle.setLeftBorderColor(sourceCellStyle.getLeftBorderColor());
        targetCellStyle.setLocked(sourceCellStyle.getLocked());
        targetCellStyle.setQuotePrefixed(sourceCellStyle.getQuotePrefixed());
        targetCellStyle.setReadingOrder(ReadingOrder.forLong(sourceCellStyle.getReadingOrder()));
        targetCellStyle.setRightBorderColor(sourceCellStyle.getRightBorderColor());
        targetCellStyle.setRotation(sourceCellStyle.getRotation());

        newCell.setCellStyle(targetCellStyle);

        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }

    }

    // 拷贝字体设置
    public static void copyFont(XSSFFont xssfFont, HSSFFont hssfFont) {
        xssfFont.setFontName(hssfFont.getFontName());
        xssfFont.setBold(hssfFont.getBold());
        xssfFont.setFontHeight(hssfFont.getFontHeight());
        xssfFont.setCharSet(hssfFont.getCharSet());
        xssfFont.setColor(hssfFont.getColor());
        xssfFont.setItalic(hssfFont.getItalic());
        xssfFont.setUnderline(hssfFont.getUnderline());
        xssfFont.setTypeOffset(hssfFont.getTypeOffset());
        xssfFont.setStrikeout(hssfFont.getStrikeout());
    }


    // 根据行列获取合并单元格
    public static CellRangeAddress getMergedRegion(List<CellRangeAddress> cellRangeAddressList, int rowNum, short cellNum) {
        for (int i = 0; i < cellRangeAddressList.size(); i++) {
            CellRangeAddress merged = cellRangeAddressList.get(i);
            if (merged.isInRange(rowNum, cellNum)) {
                // 已经获取过不再获取
                cellRangeAddressList.remove(i);
                return merged;
            }
        }
        return null;
    }

    // 拷贝图片
    public static void copyPicture(HSSFSheet source, XSSFSheet destination) {
        // 获取sheet中的图片信息
        List<Map<String, Object>> mapList = getPicturesFromHSSFSheet(source);
        XSSFDrawing drawing = destination.createDrawingPatriarch();

        for (Map<String, Object> pictureMap: mapList) {

            HSSFClientAnchor hssfClientAnchor = (HSSFClientAnchor) pictureMap.get("pictureAnchor");

            HSSFRow startRow = source.getRow(hssfClientAnchor.getRow1());
            float startRowHeight = startRow == null ? source.getDefaultRowHeightInPoints() : startRow.getHeightInPoints();

            HSSFRow endRow = source.getRow(hssfClientAnchor.getRow1());

            float endRowHeight = endRow == null ? source.getDefaultRowHeightInPoints() : endRow.getHeightInPoints();

            // hssf的单元格，每个单元格无论宽高，都被分为 宽 1024个单位 高 256个单位。
            // 32.00f 为默认的单元格单位宽度 单元格宽度 / 默认宽度 为像素宽度
            XSSFClientAnchor xssfClientAnchor = drawing.createAnchor(
                (int) (source.getColumnWidth(hssfClientAnchor.getCol1()) / 32.00f
                    / 1024 * hssfClientAnchor.getDx1() * Units.EMU_PER_PIXEL),
                (int) (startRowHeight / 256 * hssfClientAnchor.getDy1() * Units.EMU_PER_POINT),
                (int) (source.getColumnWidth(hssfClientAnchor.getCol2()) / 32.00f
                    / 1024 * hssfClientAnchor.getDx2() * Units.EMU_PER_PIXEL),
                (int) (endRowHeight / 256 * hssfClientAnchor.getDy2() * Units.EMU_PER_POINT),
                hssfClientAnchor.getCol1(),
                hssfClientAnchor.getRow1(),
                hssfClientAnchor.getCol2(),
                hssfClientAnchor.getRow2());
            xssfClientAnchor.setAnchorType(hssfClientAnchor.getAnchorType());

            drawing.createPicture(xssfClientAnchor,
                destination.getWorkbook().addPicture((byte[])pictureMap.get("pictureByteArray"),
                    Integer.parseInt(pictureMap.get("pictureType").toString())));

            System.out.println("imageInsert");
        }
    }

    /**
     * 获取图片和位置 (xls)
     */
    public static List<Map<String, Object>> getPicturesFromHSSFSheet (HSSFSheet sheet) {
        List<Map<String, Object>> mapList = new ArrayList<>();
        HSSFPatriarch drawingPatriarch = sheet.getDrawingPatriarch();
        if(drawingPatriarch==null){
            return mapList;
        }
        List<HSSFShape> list = drawingPatriarch.getChildren();
        if(list==null || list.size()==0){
            return mapList;
        }
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                Map<String, Object> map = new HashMap<>();
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = picture.getClientAnchor();
                HSSFPictureData pdata = picture.getPictureData();
                map.put("pictureAnchor", cAnchor);
                map.put("pictureByteArray", pdata.getData());
                map.put("pictureType", pdata.getPictureType());
                map.put("pictureSize", picture.getImageDimension());
                mapList.add(map);
            }
        }
        return mapList;
    }




    @Override
    public String file2PdfAndSave(String source, String target, String md5) throws Exception {

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


    public static String getPreview(String target){
        String s = target.substring(0, target.lastIndexOf(".")) + "_preview.pdf";

        return  s;
    }
}
