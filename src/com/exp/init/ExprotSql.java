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
 * @Description	����Excel��������sql
 * @author zml
 * @date 2016��5��27��
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
            int sheetCount = workBook.getNumberOfSheets();  			// Sheet������  
            for(int i=0; i<sheetCount; i++) {	
            	// �������е�Sheet
            	Sheet sheet = workBook.getSheetAt(i);
            	System.out.println("Sheet Name: " + sheet.getSheetName());
                int coordinateY = sheet.getLastRowNum();				// �õ����Y����
                for(int y = 0; y < coordinateY; y++){					// �������е���
                	Row row = sheet.getRow(y);
        			//int coordinateX = row.getPhysicalNumberOfCells();	// ��ȡ������
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
	                                fieldsName.add(Utils.getValueWithCell(cell).replace("'", ""));
	                            }
	                        }
	                        exportLog.append(new Date().toString() + ": ����" + fieldsName + "**********�����ɹ�\r\n");
	                        
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
		File file = new File("D:\\�°��Ʒ��\\DRIS.xlsx");
		ExprotSql exp = new ExprotSql();
		exp.readExcel(file);	//��ȡExcel����
		exp.createSql();		//����sql���
		exp.replaceMark();		//�滻sql����еı��
		Utils.exportSqlFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_SCRIPT, sql.toString());  	//����sql�ļ�
		Utils.exportLogFile(Constants.DEFAULT_EXPROT_DRICTORY, Constants.EXPORT_FILE_NAME_DATA_LOG, exportLog.toString());	//������־�ļ�
	}
}
