package com.exp.util;

public class Constants {
	
	// Excel表名标记
    public static final String EXCEL_TABLE_NAME_MARK = "${TableName}";

    // Excel字段描述标记
    public static final String EXCEL_FIELDS_DESCRIPTION_MARK = "${FieldsDescription}";

    // Excel字段名标记
    public static final String EXCEL_FIELDS_NAME_MARK = "${FieldsName}";
    
    // Excel数据
    public static final String EXCEL_FIELDS_DATA_MARK = "${FieldsData}";
    
    // Excel数据 datetime
    public static final String EXCEL_DATA_DATATIME_MARK = "${datetime}";
    
    // 数据库得到当前时间函数 - mysql
    public static final String DB_FUNCTION_NOW = "now()";
    
    // Sql脚本文件
    public static final String FILE_TYPE_SQL = ".sql";
    
    // 文本文件
    public static final String FILE_TYPE_TXT = ".txt";
    
    // 日志文件
    public static final String FILE_TYPE_LOG = ".log";
    
    // 默认导出路径
    public static final String DEFAULT_EXPROT_DRICTORY = "D:\\DIRS\\NEW";
    
    // 输出测试数据文件名
    public static final String EXPORT_FILE_NAME_DATA_SCRIPT = "Excel数据的Sql脚本";
    
    // 输出测试数据Log文件名
    public static final String EXPORT_FILE_NAME_DATA_LOG = "Excel生成数据日志";
    
    //日期格式化(yyyy-MM-dd HH-mm-ss-SSS)
    public static final String DATE_TIME_FORMAT_TYPE = "yyyy-MM-dd HH-mm-ss-SSS";
    
    public static final String FILE_ENCODING = "UTF-8";
    
    public static final String[] DB_DISHES_COMP_FIELDS = { "dishesCompId", "dishesCode",
        "cCode", "rawWeight", "proMethodId", "netWeight", "yieldRate", "type",
        "priority", "createTime", "creatorId", "status", "remark" };
}
