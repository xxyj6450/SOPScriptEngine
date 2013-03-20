Imports System.Data
Module modMain
    '语言类型
    Public Enum enumScriptLanguage
        TSQL = 0
        PLSQL = 1
        xQuery = 2
        JSScript = 3
        VBScript = 4
        VB = 5
        CSharp = 6
    End Enum
    '脚本类型
    Public Enum enumScriptType
        SelectScript = 0                        'Select子句,正常的Select子句,将会用SQL方式执行
        WhereScript = 1                         'Where子句
        Procedure = 2                           '存储过程
        SimpleSelectScript = 3                         '简单语句,只使用.net中Column的Expression属性中支持的语法
        SimpleWhereScript = 4
        OtherScript = 1024
    End Enum
    Public Function DataReadToDataTable(ByVal dataReader As System.Data.IDataReader) As System.Data.DataTable
        Dim dt As New System.Data.DataTable
        dt.Load(dataReader)
        Return dt
    End Function
End Module
