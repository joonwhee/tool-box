package com.joonwhee.common.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @author joonwhee
 * @date 2019/11/17
 */
public class ExcelUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 读取excel数据, 以List<String[]>形式返回数据
     *
     * @param file            excel文件流
     * @param colNum          列数
     * @param isFirstLineNeed 是否需要第一行
     * @return 以List<String[]>形式返回数据, 如果数据为空, 返回nil
     */
    public static List<String[]> readExcel(MultipartFile file, int colNum, boolean isFirstLineNeed) {
        List<String[]> result = new ArrayList<>();
        InputStream is;
        try {
            is = file.getInputStream();
            Workbook workbook = WorkbookFactory.create(is);
            // 获取excel的第一个sheet
            Sheet sheet = workbook.getSheetAt(0);
            if (null == sheet) {
                return result;
            }
            boolean isFirstLine = true;
            for (Iterator<Row> rowIterator = sheet.rowIterator(); rowIterator.hasNext(); ) {
                if (!isFirstLineNeed && isFirstLine) {
                    rowIterator.next();// 从第二行开始读
                    isFirstLine = false;
                }
                String[] valueArray = new String[colNum];
                Row row = rowIterator.next();
                for (int i = 0; i < colNum; i++) {
                    String value = "";
                    Cell cell = row.getCell(i);
                    if (cell == null) {
                        valueArray[i] = value;
                        continue;
                    }
                    int type = cell.getCellType();
                    if (type == Cell.CELL_TYPE_NUMERIC) {
                        double d = cell.getNumericCellValue();
                        BigDecimal bd = new BigDecimal(d);
                        value = bd.toString();
                    } else {
                        value = cell.getStringCellValue();
                    }
                    valueArray[i] = value;
                }
                result.add(valueArray);
            }
        } catch (IOException e) {
            LOGGER.error("读取Excel文件数据出错, IOException: ", e);
        } catch (Exception e) {
            LOGGER.error("读取Excel文件数据出错: ", e);
        }
        return result;
    }



    /**
     * 导出Excel
     *
     * @param header   标题头, 写在第一行
     * @param width    列间距
     * @param dataList 数据
     * @param out      输出流
     */
    public static void writeExcel(String[] header, int[] width, List<String[]> dataList, OutputStream out) {
        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }
        // 第一步，创建一个webbook，对应一个Excel文件
        XSSFWorkbook wb = new XSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        XSSFSheet sheet = wb.createSheet();
        // 第三步，在sheet中添加表头第0行
        XSSFRow row = sheet.createRow(0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
        XSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

        if (width.length != 0) {
            // 设置表格宽度
            for (int i = 0; i < width.length; i++) {
                // 1个汉字大约要512的宽度
                sheet.setColumnWidth(i, width[i] * 512);
            }
        }

        // 表头
        for (int i = 0; i < header.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(header[i]);
            cell.setCellStyle(style);
        }

        // 实际数据
        for (int i = 0; i < dataList.size(); i++) {
            row = sheet.createRow(i + 1);
            String[] rowValue = dataList.get(i);
            for (int j = 0; j < header.length; j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(rowValue[j]);
                cell.setCellStyle(style);
            }
        }
        try {
            wb.write(out);
        } catch (IOException e) {
            LOGGER.error("导出Excel文件数据出错", e);
        }
    }
}
