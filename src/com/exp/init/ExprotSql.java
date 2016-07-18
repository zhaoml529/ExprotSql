package com.exp.init;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.exp.util.Constants;
import com.exp.util.Utils;

/**
 * @Description	根据Excel数据生成sql
 * @author zml
 * @date 2016年5月27日
 */
public class ExprotSql {

    private String tableName = "";
    private List<String> fieldsName = new ArrayList<String>();
    private List<List<String>> dataCollection = new ArrayList<List<String>>();
    private static StringBuilder sql = new StringBuilder();
    private static StringBuilder exportLog = new StringBuilder();
    
	private void readExcel(File excelFile) {
		InputStream ins;
		try {
			ins = new FileInputStream(excelFile);
            Workbook workBook = WorkbookFactory.create(ins);
            int sheetCount = workBook.getNumberOfSheets();  			// Sheet的数量  
            for(int i=0; i<sheetCount; i++) {	
            	// 遍历所有的Sheet
            	Sheet sheet = workBook.getSheetAt(i);
            	System.out.println("Sheet Name: " + sheet.getSheetName());
                int coordinateY = sheet.getLastRowNum();				// 得到最大Y坐标
                for(int y = 0; y < coordinateY; y++){					// 遍历所有的列
                	Row row = sheet.getRow(y);
        			//int coordinateX = row.getPhysicalNumberOfCells();	// 获取总行数
        			int coordinateX = row.getLastCellNum();				// 获取总行数
        			String rowHeadValue = Utils.getValueWithCell(row.getCell(0)).replace("'", "");
        			if (row != null && !"".equals(rowHeadValue)) {
                        if (Constants.EXCEL_TABLE_NAME_MARK.equals(rowHeadValue)) {
                        	createSql();
                        	doClear();
                    		Cell cell = row.getCell(1);
                            tableName = Utils.getValueWithCell(cell).replace("'", "");
                            sql.append("-- ----------------------------\n");
                            sql.append("-- Records of " + tableName + "\n");
                            sql.append("-- ----------------------------\n");
                            exportLog.append(new Date().toString() + ": 表名" + tableName + "**********解析成功\r\n");
                            
                        } else if (Constants.EXCEL_FIELDS_NAME_MARK.equals(rowHeadValue)) {
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                                fieldsName.add(Utils.getValueWithCell(cell).replace("'", ""));
	                            }
	                        }
	                        exportLog.append(new Date().toString() + ": 列名" + fieldsName + "**********解析成功\r\n");
	                        
	                    } else if (Constants.EXCEL_FIELDS_DATA_MARK.equals(rowHeadValue)) {
	                        List<String> tmpDataList = new ArrayList<String>();
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null) {
	                                tmpDataList.add(Utils.getValueWithCell(cell));
	                            } else {
	                            	tmpDataList.add("null");
	                            }
	                        }
	                        dataCollection.add(tmpDataList);
	                        exportLog.append(new Date().toString() + ": 数据" + tmpDataList + "**********解析成功\r\n");
	                    }
        			}
                }
            } 
		} catch (Exception e) {
			e.printStackTrace();
			exportLog.append("**********资源文件解析出错**********\n" + e.toString());
            return;
		}
	}
	
	/**
     * 创建sql语句
     */
    private void createSql() {
        if (dataCollection.size() > 0 && fieldsName.size() > 0) {
            for (List<String> dataList : dataCollection) {
                sql.append("insert into " + tableName + "(");
                for (String field : fieldsName) {
                    sql.append(field + ",");
                }
                sql.delete(sql.lastIndexOf(","), sql.lastIndexOf(",") + 1);
                sql.append(") \nvalues(");
                for (String data : dataList) {
                    sql.append( data + ",");
                }
                sql.delete(sql.lastIndexOf(","), sql.lastIndexOf(",") + 1);
                sql.append(");\n");
            }
        }
    }
    
    /**
     * 替换数据中的创建时间、修改时间、创建人id、修改人id的标记
     */
    private void replaceMark() {
        String tmpSql = sql.toString().replace(Constants.EXCEL_DATA_DATATIME_MARK,
                Constants.DB_FUNCTION_NOW);
        sql = new StringBuilder(tmpSql);
    }
    
    /**
     * 清除上次的读取信息
     */
    private void doClear() {
        fieldsName.clear();
        dataCollection.clear();
    }
	
    /**
     * 入口
     */
	public static void main(String[] args) {
		File file = new File("D:\\新版菜品库\\DRIS.xlsx");
		ExprotSql exp = new ExprotSql();
		exp.readExcel(file);	//读取Excel数据
		exp.createSql();		//创建sql语句
		exp.replaceMark();		//替换sql语句中的标记
		Utils.exportSqlFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_SCRIPT, sql.toString());  	//导出sql文件
		Utils.exportLogFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_LOG, exportLog.toString());	//导出日志文件
	}
}
