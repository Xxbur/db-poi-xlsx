package com.dbPoiXlsx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class poiIntoXls {

    private final int DEFAULT_COLUMN_SIZE = 30;
    private final int flushRows = 1000;
    private Workbook writeDataWorkBook;
    private Map<String, CellStyle> cellStyleMap;
    private int currentRowNum = 0;
    private OutputStream outputStream;

    private File assertFile(String directory, String fileName) throws IOException {
        File tmpFile = new File(directory + File.separator + fileName + ".xlsx");
        if (tmpFile.exists()) {
            if (tmpFile.isDirectory()) {
                throw new IOException("File '" + tmpFile + "' exists but is a directory");
            }
            if (!tmpFile.canWrite()) {
                throw new IOException("File '" + tmpFile + "' cannot be written to");
            }
        } else {
            File parent = tmpFile.getParentFile();
            if (parent != null) {
                if (!parent.mkdirs() && !parent.isDirectory()) {
                    throw new IOException("Directory '" + parent + "' could not be created");
                }
            }
        }
        return tmpFile;
    }

    private String getCnDate(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return sdf.format(date);
    }

    public void writeExcelTitle(String directory, String fileName, String sheetName, List<String> columnNames,
                                String sheetTitle) throws Exception, IOException {
        File tmpFile = assertFile(directory, fileName);
        exportExcelTitle(tmpFile, sheetName, columnNames, sheetTitle);
        loadTplWorkbook(tmpFile);
    }

    public void writeExcelData(String directory, String fileName, String sheetName, List<List<Object>> objects)
            throws Exception {
        File tmpFile = assertFile(directory, fileName);
        outputStream = new FileOutputStream(tmpFile);
        exportExcelData(sheetName, objects);
    }

    public void dispose() throws Exception {
        try {
            if (writeDataWorkBook != null) {
                writeDataWorkBook.write(outputStream);
            }
            if (outputStream != null) {
                outputStream.flush();
                outputStream.close();
            }
            if (cellStyleMap != null) {
                cellStyleMap.clear();
            }
            cellStyleMap = null;
            outputStream = null;
            writeDataWorkBook = null;
        } catch (IOException e) {
            throw new Exception(e);
        }
    }

    private void exportExcelTitle(File file, String sheetName, List<String> columnNames,
                                  String sheetTitle) throws Exception {
        Workbook tplWorkBook = new XSSFWorkbook();
        Map<String, CellStyle> cellStyleMap = styleMap(tplWorkBook);
        CellStyle headStyle = cellStyleMap.get("head");
        Sheet sheet = tplWorkBook.getSheet(sheetName);
        if (sheet == null) {
            sheet = tplWorkBook.createSheet(sheetName);
        }
        sheet.setDefaultColumnWidth(DEFAULT_COLUMN_SIZE);
        sheet.addMergedRegion(new CellRangeAddress(currentRowNum, currentRowNum, 0, columnNames.size() - 1));
        Row rowMerged = sheet.createRow(currentRowNum);
        Cell mergedCell = rowMerged.createCell(0);
        mergedCell.setCellStyle(headStyle);
        mergedCell.setCellValue(new XSSFRichTextString());
        currentRowNum = currentRowNum + 1;
        Row row = sheet.createRow(currentRowNum);
        for (int i = 0; i < columnNames.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(headStyle);
            RichTextString text = new XSSFRichTextString(columnNames.get(i));
            cell.setCellValue(text);
        }
        currentRowNum = currentRowNum + 1;
        try {
            OutputStream ops = new FileOutputStream(file);
            tplWorkBook.write(ops);
            ops.flush();
            ops.close();
        } catch (IOException e) {
            throw new Exception(e);
        }
    }

    private void loadTplWorkbook(File file) throws Exception {
        try {
            XSSFWorkbook tplWorkBook = new XSSFWorkbook(new FileInputStream(file));
            writeDataWorkBook = new SXSSFWorkbook(tplWorkBook, flushRows);
            cellStyleMap = styleMap(writeDataWorkBook);
        } catch (IOException e) {
            throw new Exception("Excelģ���ļ�������");
        }
    }

    private void exportExcelData(String sheetName, List<List<Object>> objects) throws Exception, IOException {
        CellStyle contentStyle = cellStyleMap.get("content");
        CellStyle contentIntegerStyle = cellStyleMap.get("integer");
        CellStyle contentDoubleStyle = cellStyleMap.get("double");
        Sheet sheet = writeDataWorkBook.getSheet(sheetName);
        if (sheet == null) {
            throw new Exception("��ȡExcelģ�����");
        }
        sheet.setDefaultColumnWidth(DEFAULT_COLUMN_SIZE);
        for (List<Object> dataRow : objects) {
            Row row = sheet.createRow(currentRowNum);
            for (int j = 0; j < dataRow.size(); j++) {
                Cell contentCell = row.createCell(j);
                Object dataObject = dataRow.get(j);
                if (dataObject != null) {
                    if (dataObject instanceof Integer) {
                        contentCell.setCellStyle(contentIntegerStyle);
                        contentCell.setCellValue(Integer.parseInt(dataObject.toString()));
                    } else if (dataObject instanceof Double) {
                        contentCell.setCellStyle(contentDoubleStyle);
                        contentCell.setCellValue(Double.parseDouble(dataObject.toString()));
                    } else if (dataObject instanceof Long && dataObject.toString().length() == 13) {
                        contentCell.setCellStyle(contentStyle);
                        contentCell.setCellValue(getCnDate(new Date(Long.parseLong(dataObject.toString()))));
                    } else if (dataObject instanceof Date) {
                        contentCell.setCellStyle(contentStyle);
                        contentCell.setCellValue(getCnDate((Date) dataObject));
                    } else {
                        contentCell.setCellStyle(contentStyle);
                        contentCell.setCellValue(dataObject.toString());
                    }
                } else {
                    contentCell.setCellStyle(contentStyle);
                    contentCell.setCellValue("");
                }
            }
            currentRowNum = currentRowNum + 1;
        }
        try {
            ((SXSSFSheet) sheet).flushRows(flushRows);
        } catch (IOException e) {
            throw new Exception(e);
        }
    }

    private CellStyle createCellHeadStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        return style;
    }

    private CellStyle createCellContentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        return style;
    }

    private CellStyle createCellContent4IntegerStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        return style;
    }

    private CellStyle createCellContent4DoubleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        return style;
    }

    private Map<String, CellStyle> styleMap(Workbook workbook) {
        Map<String, CellStyle> styleMap = new LinkedHashMap<>();
        styleMap.put("head", createCellHeadStyle(workbook));
        styleMap.put("content", createCellContentStyle(workbook));
        styleMap.put("integer", createCellContent4IntegerStyle(workbook));
        styleMap.put("double", createCellContent4DoubleStyle(workbook));
        return styleMap;
    }

}
