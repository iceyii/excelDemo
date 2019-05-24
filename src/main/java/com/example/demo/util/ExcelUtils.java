package com.example.demo.util;

//import org.apache.commons.fileupload.FileItem;
//import org.apache.commons.fileupload.FileUploadException;
//import org.apache.commons.fileupload.disk.DiskFileItemFactory;
//import org.apache.commons.fileupload.servlet.ServletFileUpload;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

//import javax.servlet.http.HttpServletRequest;

/**
 * Excel导入导出工具类
 * 使用SpringBoot默认上传组件,需要配置(http:multipart:enabled: true)
 * 若未使用SpringBoot,则使用fileUpload获取上传的文件
 *
 * @author LiHuasheng
 * @version 1.1
 * @since 2019/1/26 12:55
 */
public class ExcelUtils {
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
    //读取excel的哪个sheet,默认为0
    private static final Integer SHEETINDEX = 0;
    //读取excel跳过的行数(标题,列等非数据行)
    private static final Integer TITLELINE = 0;

    /**
     * 读取上传的excel文件
     *
     * @param file      SpringBoot默认上传格式
     * @param rowMapper 列规则处理对象
     * @param <T>       对象
     * @return 返回list数据
     */
    public static <T> List<T> importExcel(MultipartFile file, ReadRowMapper<T> rowMapper) throws IOException {
        //默认列内容从excel第一行获取,如没有列名只有数据,必须定义此title
        String[] title = null;
        Workbook workBook = initWorkBook(file);
        return generateList(rowMapper, title, workBook);
    }

    /**
     * 导出excel到网络下载
     *
     * @param response  HttpServletResponse
     * @param headLine  标题行
     * @param sheetName excel下标sheet名
     * @param fileName  下载默认文件名
     * @param titles    列名数组
     * @param dataSet   传入的数据list
     * @param rowMapper 列与数据处理类
     * @throws IOException
     */
    public static void exportExcel(HttpServletResponse response, String headLine, String sheetName, String fileName, String[] titles, List<? extends Object> dataSet,
                                   WriteRowMapper rowMapper) throws IOException {
        Workbook workBook = new SXSSFWorkbook();
        OutputStream os = response.getOutputStream();
        //创建Sheet
        Sheet sheet = workBook.createSheet(sheetName);
        //设置Sheet的基本属性
        setSheet(titles, sheet);
        //生成数据
        generateContent(headLine, titles, dataSet, rowMapper, workBook, sheet);
        //导出下载
        downLoadFile(response, fileName, workBook, os);
    }

    /**
     * 设置列的一些属性
     *
     * @param titles
     * @param sheet
     */
    private static void setSheet(String[] titles, Sheet sheet) {
        //设置垂直居中
        sheet.setVerticallyCenter(true);
        //设置每列的宽
        //setCellWidth(sheet);
        //合并行,必须在创建之前,(起始行,起始行,终止列,终止列)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titles.length - 1));
        //sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, titles.length - 1));
    }

    /**
     * 生成表格数据
     *
     * @param headLine
     * @param titles
     * @param dataSet
     * @param rowMapper
     * @param workBook
     * @param sheet
     * @throws IOException
     */
    private static void generateContent(String headLine, String[] titles, List<?> dataSet, WriteRowMapper rowMapper, Workbook workBook, Sheet sheet) throws IOException {
        //生成第一行(表头)
        Row row = sheet.createRow(0);
        row.setHeight((short) 550);
        Cell headCell = row.createCell(0);
        headCell.setCellValue(headLine);
        headCell.setCellStyle(initHeadCellStyle(workBook));
        //生成第二行,列名
        Row row1 = sheet.createRow(1);
        for (int i = 0; i < titles.length; i++) {
            //列宽默认,可修改
            sheet.setColumnWidth(i, titles[i].length() * 255 * 5);
            Cell cell = row1.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(initTitleCellStyle(workBook));
        }
        //生成数据内容行
        for (int i = 0; i < dataSet.size(); i++) {
            //排除标题和列
            row = sheet.createRow(i + 2);
            List<String> values = rowMapper.handleData(dataSet.get(i));
            if (values.size() != titles.length) {
                throw new IOException("转换后的列表长度与表头数组长度不一致");
            }
            for (int j = 0; j < values.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(String.valueOf(values.get(j)));
                cell.setCellStyle(initContentCellStyle(workBook));
            }
        }
    }

    /**
     * 导出下载
     *
     * @param response
     * @param fileName
     * @param workBook
     * @param os
     * @throws UnsupportedEncodingException
     */
    private static void downLoadFile(HttpServletResponse response, String fileName, Workbook workBook, OutputStream os) throws UnsupportedEncodingException {
        if (os != null) {
            if (StringUtils.isBlank(fileName)) {
                SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddHHmm");
                fileName = formatter.format(new Date());
            }
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment; filename="
                    + URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20") + ".xlsx");
            try {
                workBook.write(os);
                os.flush();
                os.close();
            } catch (Exception e) {
//                e.printStackTrace();
                logger.info("workBook 释放失败");
            }
        }
    }

    /**
     * 处理数据生成list
     *
     * @param rowMapper
     * @param title
     * @param workBook
     * @param <T>
     * @return
     */
    private static <T> List<T> generateList(ReadRowMapper<T> rowMapper, String[] title, Workbook workBook) {
        Sheet sheet = workBook.getSheetAt(SHEETINDEX);
        List<T> list = new ArrayList<T>();
        for (Row row : sheet) {
            //跳过第一行
            if (title == null && row.getRowNum() == TITLELINE) {
                title = getRowContent(row);
                continue;
            }
            Map<String, Object> map = rowToMap(row, title);
            //去掉空行对象
            if (rowMapper.rowMap(row, map) != null) {
                list.add(rowMapper.rowMap(row, map));
            }
        }
        return list;
    }

    /**
     * 初始化WorkBook
     *
     * @param file
     * @return
     */
    private static Workbook initWorkBook(MultipartFile file) throws IOException {
        String filename = file.getOriginalFilename();
        if (suffixCheck(filename)) {
            InputStream inputStream = null;
            try {
                inputStream = file.getInputStream();
                return WorkbookFactory.create(inputStream);
            } catch (Exception e) {
//                e.printStackTrace();
                logger.info("创建Excel文件失败");
            } finally {
                if (inputStream != null) {
                    inputStream.close();
                }
            }
        } else {
            logger.info("文件类型有误");
        }
        return null;
    }

    /**
     * 验证文件结尾是否为xls或者xlsx
     *
     * @param fileName
     * @return
     */
    private static boolean suffixCheck(String fileName) {
        if (fileName.lastIndexOf(".") > -1) {
            String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
            return suffix.matches("xls|xlsx$");
        }
        return false;
    }

    /**
     * 获取每一行的内容
     *
     * @param row 列数
     * @return
     */
    private static String[] getRowContent(Row row) {
        int columnNum = row.getLastCellNum() - row.getFirstCellNum();
        String[] singleRow = new String[columnNum];
        for (int i = 0; i < columnNum; i++) {
            Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    singleRow[i] = "";
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    singleRow[i] = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    if (DateUtil.isCellDateFormatted(cell)) {
                        singleRow[i] = format.format(cell.getDateCellValue());
                    } else {
                        singleRow[i] = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    singleRow[i] = cell.getStringCellValue().trim();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    singleRow[i] = "";
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    singleRow[i] = cell.getStringCellValue();
                    if (singleRow[i] != null) {
                        singleRow[i] = singleRow[i].replaceAll("#N/A", "").trim();
                    }
                    break;
                default:
                    singleRow[i] = "";
                    break;
            }
        }
        return singleRow;
    }

    /**
     * .
     * 方法：将Excel的每行数据转换为map
     *
     * @param row
     * @param title 指定Excel的标题
     * @return
     */
    private static Map<String, Object> rowToMap(Row row, String[] title) {
        if (title == null && row.getRowNum() == 0) {
            title = getRowContent(row);
        }
        int columnNum = title.length;
        Map<String, Object> map = new HashMap<String, Object>(columnNum);
        for (int i = 0; i < columnNum; i++) {
            Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    map.put(title[i], "");
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    map.put(title[i], cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        map.put(title[i], cell.getDateCellValue());
                    } else {
                        map.put(title[i], cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    map.put(title[i], cell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    map.put(title[i], "");
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    map.put(title[i], cell.getStringCellValue());
                    break;
                default:
                    map.put(title[i], "");
                    break;
            }
        }
        return map;
    }

    /**
     * 初始化头部标题样式(可自定义修改)
     */
    private static CellStyle initHeadCellStyle(Workbook workBook) {
        CellStyle headerStyle = workBook.createCellStyle();// 创建标题样式
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // 设置垂直居中
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER); // 设置水平居中
        Font headerFont = workBook.createFont(); // 创建字体样式
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
        headerFont.setFontName("微软雅黑"); // 设置字体类型
        headerFont.setFontHeightInPoints((short) 14); // 设置字体大小
        headerStyle.setFont(headerFont); // 为标题样式设置字体样式
        return headerStyle;
    }

    /**
     * 初始化列样式(可自定义修改)
     */
    private static CellStyle initTitleCellStyle(Workbook workBook) {
        CellStyle headerStyle = workBook.createCellStyle();// 创建标题样式
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // 设置垂直居中
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER); // 设置水平居中
        headerStyle.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
        headerStyle.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
        headerStyle.setBorderTop(CellStyle.BORDER_THIN);// 上边框
        headerStyle.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        Font headerFont = workBook.createFont(); // 创建字体样式
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
        headerFont.setFontName("微软雅黑"); // 设置字体类型
        headerFont.setFontHeightInPoints((short) 10); // 设置字体大小
        headerStyle.setFont(headerFont); // 为标题样式设置字体样式
        return headerStyle;
    }

    /**
     * 初始化excel内容表格样式(可自定义修改)
     */
    private static CellStyle initContentCellStyle(Workbook workBook) {
        CellStyle cell_Style = workBook.createCellStyle();// 设置字体样式
        cell_Style.setAlignment(CellStyle.ALIGN_CENTER);
        cell_Style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 垂直对齐居中
        cell_Style.setWrapText(true); // 设置为自动换行
        Font cell_Font = workBook.createFont();
        cell_Font.setFontName("微软雅黑");
        cell_Font.setFontHeightInPoints((short) 10);
        cell_Style.setFont(cell_Font);
        cell_Style.setBorderBottom(CellStyle.BORDER_THIN); // 下边框
        cell_Style.setBorderLeft(CellStyle.BORDER_THIN);// 左边框
        cell_Style.setBorderTop(CellStyle.BORDER_THIN);// 上边框
        cell_Style.setBorderRight(CellStyle.BORDER_THIN);// 右边框
        return cell_Style;
    }

    /**
     * 设置列宽
     *
     * @param sheet
     */
    private static void setCellWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 4400);
        sheet.setColumnWidth(1, 11000);
        sheet.setColumnWidth(2, 2400);
        sheet.setColumnWidth(3, 5000);
        sheet.setColumnWidth(4, 5000);
        sheet.setColumnWidth(5, 4400);
        sheet.setColumnWidth(6, 3300);
    }
//---------------------------------------------------------非springboot使用以下方法
//    /**
//     * 使用fileupload读取上传的文件生成list
//     * @param request
//     * @param rowMapper
//     * @param <T>
//     * @return
//     */
//    public static <T> List<T> importExcelFromFileUpload(HttpServletRequest request, ReadRowMapper<T> rowMapper) {
//        Workbook workBook = initWorkBookFromFileUpload(request);
//        //默认列内容从excel第一行获取,如没有列名只有数据,必须定义此title
//        String[] title = null;
//        return generateList(rowMapper, title, workBook);
//    }
//    /**
//     * 从网络上传中读取excel(fileupload工具)
//     *
//     * @return
//     * @throws EncryptedDocumentException
//     * @throws InvalidFormatException
//     * @throws IOException
//     */
//    private static Workbook initWorkBookFromFileUpload(HttpServletRequest request){
//        FileItem item;
//        try {
//            item = getFileItem(request);
//            String filename = item.getName();
//            if (suffixCheck(filename)) {
//                return WorkbookFactory.create(item.getInputStream());
//            }
//        } catch (FileUploadException | IOException | InvalidFormatException e) {
//            e.printStackTrace();
//        }
//        return null;
//    }
//    /**
//     * 使用file-upload抓取上传的文件
//     *
//     * @param request
//     * @return
//     * @throws FileUploadException
//     */
//    private static FileItem getFileItem(HttpServletRequest request) throws FileUploadException {
//        DiskFileItemFactory fileItemFactory = new DiskFileItemFactory();
//        ServletFileUpload sfu = new ServletFileUpload(fileItemFactory);
//        List<FileItem> items = sfu.parseRequest(request);
//        return items.get(0);
//    }
}