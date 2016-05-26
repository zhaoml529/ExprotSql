package com.exp.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public class Utils {
	/**
     * 根据Excel每个单元格的类型不同得到单元格的Value
     * @param cell
     * @return
     */
    public static String getValueWithCell(Cell cell) {
        Object obj = null;
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_STRING:
            obj = cell.getStringCellValue();
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            obj = cell.getBooleanCellValue();
            break;
        case Cell.CELL_TYPE_FORMULA:
            obj = cell.getCellFormula();
            break;
        case Cell.CELL_TYPE_ERROR:
            obj = cell.getErrorCellValue();
            break;
        case Cell.CELL_TYPE_BLANK:
            obj = "";
            break;
        case Cell.CELL_TYPE_NUMERIC:
            obj = (int)cell.getNumericCellValue();
            break; 
        }
        return obj == null ? null : "'" + obj.toString() + "'";
    }
    
    public static void exportFile(String exportPath, String filename, String data, String fileType) {
        OutputStream ops = null;
        try {
            SimpleDateFormat format = new SimpleDateFormat(Constants.DATE_TIME_FORMAT_TYPE);
            File logDir = new File(exportPath);
            if (!logDir.exists()) {
                logDir.mkdir();
            }
            File logFile = new File(exportPath + "\\"
                    + format.format(new Date()) + filename + fileType);
            ops = new FileOutputStream(logFile);
            ops.write(data.getBytes(Constants.FILE_ENCODING));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                ops.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
