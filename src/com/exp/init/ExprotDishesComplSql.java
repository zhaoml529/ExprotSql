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
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
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
	                            	if(x == 3 || x == 6) {
	                            		FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator(); 
	                            		CellValue cellValue = evaluator.evaluate(cell);
	                            		tmpDataList.add("'"+String.valueOf(new Double(cellValue.getNumberValue()).intValue())+"'");
	                            	} else {
	                            		tmpDataList.add(Utils.getValueWithCell(cell));
	                            	}
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
            	StringBuffer begin = new StringBuffer();
            	begin.append("insert into " + tableName + "(");
                for (String key : Constants.DB_DISHES_COMP_FIELDS) {
                    begin.append(key + ",");
                }
                begin.delete(begin.lastIndexOf(","), begin.lastIndexOf(",") + 1);
                begin.append(") \nvalues(");
                begin.append("null," + dataList.get(1) + "," + dataList.get(2) + "," );
                
                sql.append(begin);
                sql.append(dataList.get(4) + "," + dataList.get(5) + "," + dataList.get(6) + "," + dataList.get(7) + ",1,null,null,0,1,"+dataList.get(38)+");\n");
                
                if(!"null".equals(dataList.get(8))){	//第一个辅料为空，后面的辅料都不用添加了
                	sql.append(begin);
                	sql.append(dataList.get(9) + "," + dataList.get(5) + "," + dataList.get(11) + "," + dataList.get(12) + ",1,null,null,0,2,"+dataList.get(38)+");\n");
                	if(!"null".equals(dataList.get(13))){
                		sql.append(begin);
                		sql.append(dataList.get(14) + "," + dataList.get(5) + "," + dataList.get(16) + "," + dataList.get(17) + ",1,null,null,0,2,"+dataList.get(38)+");\n");
                		if(!"null".equals(dataList.get(18))){
                			sql.append(begin);
                			sql.append(dataList.get(19) + "," + dataList.get(5) + "," + dataList.get(21) + "," + dataList.get(22) + ",1,null,null,0,2,"+dataList.get(38)+");\n");
                			if(!"null".equals(dataList.get(23))){
                				sql.append(begin);
                				sql.append(dataList.get(24) + "," + dataList.get(5) + "," + dataList.get(26) + "," + dataList.get(27) + ",1,null,null,0,2,"+dataList.get(38)+");\n");
                			}
                		}
                	}
                }
                
                if(!"null".equals(dataList.get(28))){	//第一个调料为空，后面的调料都不用添加了
                	sql.append(begin);
                	sql.append(dataList.get(29) + "," + dataList.get(5) + ",null,null,3,null,null,0,1,"+dataList.get(38)+");\n");
                	if(!"null".equals(dataList.get(30))){
                		sql.append(begin);
                		sql.append(dataList.get(31) + "," + dataList.get(5) + ",null,null,3,null,null,0,1,"+dataList.get(38)+");\n");
                		if(!"null".equals(dataList.get(32))){
                			sql.append(begin);
                			sql.append(dataList.get(33) + "," + dataList.get(5) + ",null,null,3,null,null,0,1,"+dataList.get(38)+");\n");
                			if(!"null".equals(dataList.get(34))){
                				sql.append(begin);
                				sql.append(dataList.get(35) + "," + dataList.get(5) + ",null,null,3,null,null,0,1,"+dataList.get(38)+");\n");
                				if(!"null".equals(dataList.get(36))){
                					sql.append(begin);
                					sql.append(dataList.get(37) + "," + dataList.get(5) + ",null,null,3,null,null,0,1,"+dataList.get(38)+");\n");
                				}
                			}
                		}
                	}
                	
                }
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
		File file = new File("D:\\菜品原料清单.xlsx");
		ExprotDishesComplSql exp = new ExprotDishesComplSql();
		exp.readExcel(file);	//读取Excel数据
		exp.createSql();		//创建sql语句
		exp.replaceMark();		//替换sql语句中的标记
		Utils.exportSqlFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_SCRIPT, sql.toString());  	//导出sql文件
		Utils.exportLogFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_LOG, exportLog.toString());	//导出日志文件
	}
}
