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
    private List<List<String>> dataCollection = new ArrayList<List<String>>(3050);
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
	                        List<String> tmpDataList = new ArrayList<String>(60);
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                            	/*if(x == 3 || x == 6 || x == 9) {
	                            		FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator(); 
	                            		CellValue cellValue = evaluator.evaluate(cell);
	                            		tmpDataList.add("'"+String.valueOf(new Double(cellValue.getNumberValue()).intValue())+"'");
	                            	} else {
	                            		tmpDataList.add(Utils.getValueWithCell(cell));
	                            	}*/
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
    
    /**
     * 创建sql语句
     */
    private void createSql() {
        if (dataCollection.size() > 0 && fieldsName.size() > 0) {
            for (List<String> dataList : dataCollection) {
            	int i = 1;
            	StringBuffer begin = new StringBuffer();
            	begin.append("insert into " + tableName + "(");
                for (String key : Constants.DB_DISHES_COMP_FIELDS) {
                    begin.append(key + ",");
                }
                begin.delete(begin.lastIndexOf(","), begin.lastIndexOf(",") + 1);
                begin.append(") \nvalues(");
                begin.append("null," + dataList.get(2) + ",");
                System.out.println(dataList.get(1)+" - "+dataList.get(2));
                sql.append(begin);
                sql.append(dataList.get(3) + "," + dataList.get(5) + "," + dataList.get(6) + "," + dataList.get(8) + "," + dataList.get(9) + ",1,"+ (i++) +",null,0,1,null);\n");
                
                if(!"null".equals(dataList.get(10)) && !"".equals(dataList.get(10))){	//第一个辅料为空，后面的辅料都不用添加了
                	sql.append(begin);
                	sql.append(dataList.get(10) + "," + dataList.get(12) + "," + dataList.get(13) + "," + dataList.get(15) + "," + dataList.get(16) + ",2,"+ (i++) +",null,0,1,null);\n");
                	if(!"null".equals(dataList.get(17)) && !"".equals(dataList.get(17))){	//第二个辅料
                		sql.append(begin);
                		sql.append(dataList.get(17) + "," + dataList.get(19) + "," + dataList.get(20) + "," + dataList.get(22) + "," + dataList.get(23) + ",2,"+ (i++) +",null,0,1,null);\n");
                		if(!"null".equals(dataList.get(24)) && !"".equals(dataList.get(24))){	//第三个辅料
                			sql.append(begin);
                			sql.append(dataList.get(24) + "," + dataList.get(26) + "," + dataList.get(27) + "," + dataList.get(29) + "," + dataList.get(30) + ",2,"+ (i++) +",null,0,1,null);\n");
                			if(!"null".equals(dataList.get(31)) && !"".equals(dataList.get(31))){	//第四个辅料
                				sql.append(begin);
                				sql.append(dataList.get(31) + "," + dataList.get(33) + "," + dataList.get(34) + "," + dataList.get(36) + "," + dataList.get(37) + ",2,"+ (i++) +",null,0,1,null);\n");
                			}
                		}
                	}
                }
                
                if(!"null".equals(dataList.get(38)) && !"".equals(dataList.get(38))){	//第一个调料为空，后面的调料都不用添加了
                	sql.append(begin);
                	sql.append(dataList.get(38) + "," + dataList.get(40) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                	if(!"null".equals(dataList.get(41)) && !"".equals(dataList.get(41))){	//第二个调料
                		sql.append(begin);
                		sql.append(dataList.get(41) + "," + dataList.get(43) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                		if(!"null".equals(dataList.get(44)) && !"".equals(dataList.get(44))){	//第三个调料
                			sql.append(begin);
                			sql.append(dataList.get(44) + "," + dataList.get(46) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                			if(!"null".equals(dataList.get(47)) && !"".equals(dataList.get(47))){	//第四个调料
                				sql.append(begin);
                				sql.append(dataList.get(47) + "," + dataList.get(49) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                				if(!"null".equals(dataList.get(50)) && !"".equals(dataList.get(50))){	//第五个调料
                					sql.append(begin);
                					sql.append(dataList.get(50) + "," + dataList.get(52) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                					if(!"null".equals(dataList.get(53)) && !"".equals(dataList.get(53))){	//第六个调料
                    					sql.append(begin);
                    					sql.append(dataList.get(53) + "," + dataList.get(55) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                    					if(!"null".equals(dataList.get(56)) && !"".equals(dataList.get(56))){	//第七个调料
                        					sql.append(begin);
                        					sql.append(dataList.get(56) + "," + dataList.get(58) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                        				}
                    				}
                				}
                			}
                		}
                	}
                }
                if(!"null".equals(dataList.get(59)) && !"".equals(dataList.get(59))){	//食用油
                	sql.append(begin);
                	sql.append(dataList.get(59) + "," + dataList.get(61) + ",null,null,null,4,"+ (i++) +",null,0,1,null);\n");
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
		File file = new File("D:\\新版菜品库0803\\西安\\菜品库原料清单(无公式).xlsx");
		ExprotDishesComplSql exp = new ExprotDishesComplSql();
		exp.readExcel(file);	//读取Excel数据
		exp.createSql();		//创建sql语句
		exp.replaceMark();		//替换sql语句中的标记
		Utils.exportSqlFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_SCRIPT, sql.toString());  	//导出sql文件
		Utils.exportLogFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_LOG, exportLog.toString());	//导出日志文件
	}
}
