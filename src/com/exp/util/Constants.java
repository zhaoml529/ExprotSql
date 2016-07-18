package com.exp.util;

public class Constants {
	
	// Excel�������
    public static final String EXCEL_TABLE_NAME_MARK = "${TableName}";

    // Excel�ֶ��������
    public static final String EXCEL_FIELDS_DESCRIPTION_MARK = "${FieldsDescription}";

    // Excel�ֶ������
    public static final String EXCEL_FIELDS_NAME_MARK = "${FieldsName}";
    
    // Excel����
    public static final String EXCEL_FIELDS_DATA_MARK = "${FieldsData}";
    
    // Excel���� datetime
    public static final String EXCEL_DATA_DATATIME_MARK = "${datetime}";
    
    // ���ݿ�õ���ǰʱ�亯�� - mysql
    public static final String DB_FUNCTION_NOW = "now()";
    
    // Sql�ű��ļ�
    public static final String FILE_TYPE_SQL = ".sql";
    
    // �ı��ļ�
    public static final String FILE_TYPE_TXT = ".txt";
    
    // ��־�ļ�
    public static final String FILE_TYPE_LOG = ".log";
    
    // Ĭ�ϵ���·��
    public static final String DEFAULT_EXPROT_DRICTORY = "D:\\DIRS\\NEW";
    
    // ������������ļ���
    public static final String EXPORT_FILE_NAME_DATA_SCRIPT = "Excel���ݵ�Sql�ű�";
    
    // �����������Log�ļ���
    public static final String EXPORT_FILE_NAME_DATA_LOG = "Excel����������־";
    
    //���ڸ�ʽ��(yyyy-MM-dd HH-mm-ss-SSS)
    public static final String DATE_TIME_FORMAT_TYPE = "yyyy-MM-dd HH-mm-ss-SSS";
    
    public static final String FILE_ENCODING = "UTF-8";
    
    public static final String[] DB_DISHES_COMP_FIELDS = { "dishesCompId", "dishesCode",
        "cCode", "rawWeight", "proMethodId", "netWeight", "yieldRate", "type",
        "priority", "createTime", "creatorId", "status", "remark" };
}
