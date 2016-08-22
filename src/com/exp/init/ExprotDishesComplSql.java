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
 * @Description ������Ʒ��ʳ�ģ�ԭ�ϣ�����
 * @author zhaomingliang
 * @date 2016��5��30��
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
            int sheetCount = workBook.getNumberOfSheets();  			// Sheet������  
            for(int i=0; i<sheetCount; i++) {							// �������е�Sheet
            	Sheet sheet = workBook.getSheetAt(i);
            	System.out.println("Sheet Name: " + sheet.getSheetName());
                int coordinateY = sheet.getLastRowNum();				// �õ����Y����
                for(int y = 0; y < coordinateY; y++){					// �������е���
                	Row row = sheet.getRow(y);
        			int coordinateX = row.getLastCellNum();				// ��ȡ������
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
                            exportLog.append(new Date().toString() + ": ����" + tableName + "**********�����ɹ�\r\n");
                            
                        } else if (Constants.EXCEL_FIELDS_NAME_MARK.equals(rowHeadValue)) {
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                            	String fieldName = Utils.getValueWithCell(cell).replace("'", ""); 
	                                fieldsName.put(fieldName, fieldName);
	                            }
	                        }
	                        exportLog.append(new Date().toString() + ": ����" + fieldsName + "**********�����ɹ�\r\n");
	                        
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
	                        exportLog.append(new Date().toString() + ": ����" + tmpDataList + "**********�����ɹ�\r\n");
	                    }
        			}
                }
            }
		} catch (Exception e) {
			e.printStackTrace();
			exportLog.append("**********��Դ�ļ���������**********\n" + e.toString());
            return;
		}
    }
    
    /**
     * ����sql���
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
                
                if(!"null".equals(dataList.get(10)) && !"".equals(dataList.get(10))){	//��һ������Ϊ�գ�����ĸ��϶����������
                	sql.append(begin);
                	sql.append(dataList.get(10) + "," + dataList.get(12) + "," + dataList.get(13) + "," + dataList.get(15) + "," + dataList.get(16) + ",2,"+ (i++) +",null,0,1,null);\n");
                	if(!"null".equals(dataList.get(17)) && !"".equals(dataList.get(17))){	//�ڶ�������
                		sql.append(begin);
                		sql.append(dataList.get(17) + "," + dataList.get(19) + "," + dataList.get(20) + "," + dataList.get(22) + "," + dataList.get(23) + ",2,"+ (i++) +",null,0,1,null);\n");
                		if(!"null".equals(dataList.get(24)) && !"".equals(dataList.get(24))){	//����������
                			sql.append(begin);
                			sql.append(dataList.get(24) + "," + dataList.get(26) + "," + dataList.get(27) + "," + dataList.get(29) + "," + dataList.get(30) + ",2,"+ (i++) +",null,0,1,null);\n");
                			if(!"null".equals(dataList.get(31)) && !"".equals(dataList.get(31))){	//���ĸ�����
                				sql.append(begin);
                				sql.append(dataList.get(31) + "," + dataList.get(33) + "," + dataList.get(34) + "," + dataList.get(36) + "," + dataList.get(37) + ",2,"+ (i++) +",null,0,1,null);\n");
                			}
                		}
                	}
                }
                
                if(!"null".equals(dataList.get(38)) && !"".equals(dataList.get(38))){	//��һ������Ϊ�գ�����ĵ��϶����������
                	sql.append(begin);
                	sql.append(dataList.get(38) + "," + dataList.get(40) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                	if(!"null".equals(dataList.get(41)) && !"".equals(dataList.get(41))){	//�ڶ�������
                		sql.append(begin);
                		sql.append(dataList.get(41) + "," + dataList.get(43) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                		if(!"null".equals(dataList.get(44)) && !"".equals(dataList.get(44))){	//����������
                			sql.append(begin);
                			sql.append(dataList.get(44) + "," + dataList.get(46) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                			if(!"null".equals(dataList.get(47)) && !"".equals(dataList.get(47))){	//���ĸ�����
                				sql.append(begin);
                				sql.append(dataList.get(47) + "," + dataList.get(49) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                				if(!"null".equals(dataList.get(50)) && !"".equals(dataList.get(50))){	//���������
                					sql.append(begin);
                					sql.append(dataList.get(50) + "," + dataList.get(52) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                					if(!"null".equals(dataList.get(53)) && !"".equals(dataList.get(53))){	//����������
                    					sql.append(begin);
                    					sql.append(dataList.get(53) + "," + dataList.get(55) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                    					if(!"null".equals(dataList.get(56)) && !"".equals(dataList.get(56))){	//���߸�����
                        					sql.append(begin);
                        					sql.append(dataList.get(56) + "," + dataList.get(58) + ",null,null,null,3,"+ (i++) +",null,0,1,null);\n");
                        				}
                    				}
                				}
                			}
                		}
                	}
                }
                if(!"null".equals(dataList.get(59)) && !"".equals(dataList.get(59))){	//ʳ����
                	sql.append(begin);
                	sql.append(dataList.get(59) + "," + dataList.get(61) + ",null,null,null,4,"+ (i++) +",null,0,1,null);\n");
                }
            }
        }
    }
    
    /**
     * �滻�����еĴ���ʱ�䡢�޸�ʱ�䡢������id���޸���id�ı��
     */
    private void replaceMark() {
        String tmpSql = sql.toString().replace(Constants.EXCEL_DATA_DATATIME_MARK,
                Constants.DB_FUNCTION_NOW);
        sql = new StringBuilder(tmpSql);
    }
    
    /**
     * ����ϴεĶ�ȡ��Ϣ
     */
    private void doClear() {
        fieldsName.clear();
        dataCollection.clear();
    }
    
    /**
     * ���
     */
	public static void main(String[] args) {
		File file = new File("D:\\�°��Ʒ��0803\\����\\��Ʒ��ԭ���嵥(�޹�ʽ).xlsx");
		ExprotDishesComplSql exp = new ExprotDishesComplSql();
		exp.readExcel(file);	//��ȡExcel����
		exp.createSql();		//����sql���
		exp.replaceMark();		//�滻sql����еı��
		Utils.exportSqlFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_SCRIPT, sql.toString());  	//����sql�ļ�
		Utils.exportLogFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_LOG, exportLog.toString());	//������־�ļ�
	}
}
