package com.exp.init;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.exp.util.Constants;
import com.exp.util.Utils;

/**
 * @Description 导出菜品的食材（原料）数据
 * @author zhaomingliang
 * @date 2016年5月30日
 */
public class ExprotDishesComplSql {
	
	private String tableName = "";
    private Map<String, String> fieldsName = new HashMap<String, String>();
    private List<List<String>> dataCollection = new ArrayList<List<String>>();
    private static StringBuilder sql = new StringBuilder();
    private static StringBuilder exportLog = new StringBuilder();
    
    private void readExcel(File excelFile) {
    	InputStream ins;
		try {
			ins = new FileInputStream(excelFile);
            Workbook workBook = WorkbookFactory.create(ins);
            int sheetCount = workBook.getNumberOfSheets();  			// Sheet的数量  
            for(int i=0; i<sheetCount; i++) {							// 遍历所有的Sheet
            	Sheet sheet = workBook.getSheetAt(i);
            	System.out.println("Sheet Name: " + sheet.getSheetName());
                int coordinateY = sheet.getLastRowNum();				// 得到最大Y坐标
                for(int y = 0; y < coordinateY; y++){					// 遍历所有的列
                	Row row = sheet.getRow(y);
        			int coordinateX = row.getLastCellNum();				// 获取总行数
        			String rowHeadValue = Utils.getValueWithCell(row.getCell(0)).replace("'", "");
        			if (row != null && !"".equals(rowHeadValue)) {
                        if (Constants.EXCEL_TABLE_NAME_MARK.equals(rowHeadValue)) {
                        	//createSql();
                        	//doClear();
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
	                            	String fieldName = Utils.getValueWithCell(cell).replace("'", ""); 
	                                fieldsName.put(fieldName, fieldName);
	                            }
	                        }
	                        exportLog.append(new Date().toString() + ": 列名" + fieldsName + "**********解析成功\r\n");
	                        
	                    } else if (Constants.EXCEL_FIELDS_DATA_MARK.equals(rowHeadValue)) {
	                        List<String> tmpDataList = new ArrayList<String>();
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                                tmpDataList.add(Utils.getValueWithCell(cell));
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
}
