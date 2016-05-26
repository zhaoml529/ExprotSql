package com.exp.init;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.exp.util.Constants;
import com.exp.util.Utils;

public class ExprotSql {

    private String tableName = null;
    private List<String> fieldsName = new ArrayList<String>();
    private List<List<String>> testData = new ArrayList<List<String>>();
    private StringBuilder sql = new StringBuilder();
    
	private void readExcel(File excelFile) {
		InputStream ins;
		try {
			ins = new FileInputStream(excelFile);
            Workbook workBook = WorkbookFactory.create(ins);
            int sheetCount = workBook.getNumberOfSheets();  //Sheet������  
            for(int i=0; i<sheetCount; i++) {				//�������е�Sheet
            	Sheet sheet = workBook.getSheetAt(i);
            	System.out.println("Sheet Name: " + sheet.getSheetName());
            	// �õ����Y����
                int coordinateY = sheet.getPhysicalNumberOfRows();
                for(int y = 0; y < coordinateY; y++){		//�������е���
                	Row row = sheet.getRow(y);
        			int coordinateX = row.getPhysicalNumberOfCells();	//��ȡ������
        			String rowHeadValue = Utils.getValueWithCell(row.getCell(0)).replace("'", "");
        			if (row != null && !"".equals(rowHeadValue)) {
                        if (Constants.EXCEL_TABLE_NAME_MARK.equals(rowHeadValue)) {
                    		Cell cell = row.getCell(1);
                            tableName = Utils.getValueWithCell(cell).replace("'", "");
                            System.out.println("����" + tableName + "**********�����ɹ�\n");
                        } else if (Constants.EXCEL_FIELDS_NAME_MARK.equals(rowHeadValue)) {
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                            	System.out.println(Utils.getValueWithCell(cell).replace("'", ""));
	                                fieldsName.add(Utils.getValueWithCell(cell).replace("'", ""));
	                            }
	                        }
	                        System.out.println("����" + fieldsName + "**********�����ɹ�\n");
	                    } else if (Constants.EXCEL_TEST_DATA_MARK.equals(rowHeadValue)) {
	                        List<String> tmpDataList = new ArrayList<String>();
	                        for (int x = 1; x < coordinateX; x++) {
	                        	Cell cell = row.getCell(x);
	                            if(cell != null){
	                                tmpDataList.add(Utils.getValueWithCell(cell));
	                            }
	                        }
	                        testData.add(tmpDataList);
	                        System.out.println("����" + testData + "**********�����ɹ�\n");
	                    }
        			}
                }
            } 
		} catch (Exception e) {
			e.printStackTrace();
            System.out.println("**********��Դ�ļ���������**********\n" + e.toString());
            return;
		}
	}
	
	/**
     * ����sql���
     */
    private void createSql(){
        if (testData.size() > 0 && fieldsName.size() > 0) {
            for (List<String> testDataList : testData) {
                sql.append("delete from " + tableName + " where " + fieldsName.get(0) + " = " + testDataList.get(0) + "\n");
                sql.append("insert into " + tableName + "(");
                for (String field : fieldsName) {
                    sql.append(field + ",");
                }
                sql.delete(sql.lastIndexOf(","), sql.lastIndexOf(",") + 1);
                sql.append(") \nvalues(");
                for (String data : testDataList) {
                    sql.append( data + ",");
                }
                sql.delete(sql.lastIndexOf(","), sql.lastIndexOf(",") + 1);
                sql.append(")\n");
            }
        }
    }
    
    /**
     * �滻�����е��޸�ʱ�䡢����ʱ�䡢�޸��������������������ı��
     */
    private String replaceMark(){
        String tmpSql = sql.toString().replace(Constants.EXCEL_TEST_DATA_DATATIME_MARK,
                Constants.DB_FUNCTION_NOW);
        sql = new StringBuilder(tmpSql);
        return sql.toString();
    }
	
	public static void main(String[] args) {
		File file = new File("D:\\DRIS.xlsx");
		ExprotSql exp = new ExprotSql();
		exp.readExcel(file);
		exp.createSql();
		String sql = exp.replaceMark();
		Utils.exportFile("D:\\", Constants.EXPORT_FILE_NAME_TEST_DATA_SCRIPT, sql, Constants.FILE_TYPE_SQL);
	}
}
