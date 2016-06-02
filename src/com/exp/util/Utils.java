package com.exp.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
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
            obj = "null";
            break;
        case Cell.CELL_TYPE_NUMERIC:
        	DecimalFormat df = new DecimalFormat("#.#####");
        	obj = df.format(cell.getNumericCellValue()); 
            break; 
        }
        return obj == null ? null : "null".equals(obj.toString()) ? obj.toString() : "'" + obj.toString() + "'";
    }
    
    /**
     * 输出程序运行的跟踪日志
     * 
     * @param exportPath:日志存放的路径
     * @param filename:输出日志文件的文件名
     * @param exportLog:日志文件的内容
     */
    public static void exportLogFile(String exportPath, String filename, String exportLog) {
        exportFile(exportPath, filename, exportLog, Constants.FILE_TYPE_LOG);
    }

    /**
     * 输出脚本文件
     * 
     * @param exportPath:脚本存放的路径
     * @param filename:输出脚本文件的文件名
     * @param sql:脚本文件的内容
     */
    public static void exportSqlFile(String exportPath, String filename, String sql) {
        exportFile(exportPath, filename, sql, Constants.FILE_TYPE_SQL);
    }
    
    public static void exportFile(String exportPath, String filename, String data, String fileType) {
        OutputStream ops = null;
        try {
            SimpleDateFormat format = new SimpleDateFormat(Constants.DATE_TIME_FORMAT_TYPE);
            File logDir = new File(exportPath);
            if (!logDir.exists()) {
                logDir.mkdir();
            }
            File logFile = new File(exportPath + "\\" + format.format(new Date()) + filename + fileType);
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
