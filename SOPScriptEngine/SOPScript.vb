Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.SqlServer.Server
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
<Microsoft.VisualBasic.ComClass()> <System.Serializable()> Partial Public Class ScriptEngine


    '计算一个表达式,获取返回值 compute an expression and get the value
    '参数:
    'Language:脚本语言,默认为T-SQL
    'Expression:表达式
    'DataSource:数据源
    'ExtendDataSource扩展数据源
    'Fields:字段对应表
    'AutoAddQuotes是否自动添加单引号
    '返回值:返回一个对象,对应SQL SERVER的sql_variant类型
    <Microsoft.SqlServer.Server.SqlFunction(DataAccess:=DataAccessKind.Read, SystemDataAccess:=SystemDataAccessKind.Read)> _
    Public Shared Function ExecuteScalar(ByVal Language As SqlInt16, ByVal Expression As SqlString, _
                                         ByVal DataSource As SqlString, ByVal Filter As SqlString, ByVal MaxRows As SqlInt32, ByVal Fields As SqlString, _
                                                 ByVal AutoAddQuotes As SqlBoolean) As Object
        Dim _Datasource As DataSet
        Dim Scripts As ScriptTree
        If Trim(Expression.ToString) = "" Then Return ""
        '_Datasource = New System.Data.DataSet
        Using Connection = New System.Data.SqlClient.SqlConnection("context connection=true")
            '取得记录集
            _Datasource = getDataSource(Connection, DataSource.Value, MaxRows.Value, Fields.Value)
            _Datasource.ExtendedProperties("Filter") = Filter.Value
            '替换关键字
            Expression = New KeywordsProcessor(Expression.Value, _Datasource, Filter.Value, AutoAddQuotes.Value).Expression()
            '分割子表达式
            'Scripts = SplitExpression(Expression.Value)
            '创建脚本树
            Scripts = New ScriptTree(Expression.Value, _Datasource)

            Dim retobj As System.Data.DataRow
            '从脚本树取得值
            retobj = Scripts.getValue(Connection)
            If Not (retobj Is Nothing) Then
                Return retobj.Item(0)
            Else
                Return retobj
            End If
        End Using
        Scripts = Nothing
    End Function

    <Microsoft.SqlServer.Server.SqlFunction(TableDefinition:="data1 sql_variant,data2 sql_variant,data3 sql_variant,data4 sql_variant,data5 sql_variant, " & _
                                            " data6 sql_variant,data7 sql_variant,data8 sql_variant,data9 sql_variant,data10 sql_variant," & _
                                            " data11 sql_variant,data12 sql_variant,data13 sql_variant,data14 sql_variant,data15 sql_variant, " & _
                                            " data16 sql_variant,data17 sql_variant,data18 sql_variant,data19 sql_variant,data20 sql_variant" _
                                            , FillRowMethodName:="FillRows", Name:="ExecuteTable", DataAccess:=DataAccessKind.Read, SystemDataAccess:=SystemDataAccessKind.Read)> _
    Public Shared Function ExecuteTable(ByVal Language As SqlInt16, ByVal Expression As SqlString, _
                                         ByVal DataSource As SqlString, ByVal Filter As SqlString, ByVal MaxRows As SqlInt32, ByVal Fields As SqlString, _
                                                 ByVal AutoAddQuotes As SqlBoolean) As IEnumerable
        Dim _Datasource As DataSet
        Dim Scripts As ScriptTree
        Dim dataTable As New System.Data.DataTable
        If Expression = "" Then Return ""
        '_Datasource = New System.Data.DataSet
        Using Connection = New System.Data.SqlClient.SqlConnection("context connection=true")
            '取得记录集
            _Datasource = getDataSource(Connection, DataSource.Value, MaxRows.Value, Fields.Value)
            _Datasource.ExtendedProperties("Filter") = Filter.Value
            '替换关键字
            Expression = New KeywordsProcessor(Expression.Value, _Datasource, Filter.Value, AutoAddQuotes.Value).Expression()
            '分割子表达式
            'Scripts = SplitExpression(Expression.Value)
            '创建脚本树
            Scripts = New ScriptTree(Expression.Value, _Datasource)

            Dim retobj As System.Data.DataRow
            '从脚本树取得值
            'get value from the script tree
            retobj = Scripts.getValue(Connection)
            dataTable = retobj.Table.Copy
            Return dataTable.Rows
            Scripts = Nothing
        End Using

    End Function
    Private Shared Sub FillRows(ByVal Row As Object, <Out()> ByRef data1 As Object, <Out()> ByRef data2 As Object, <Out()> ByRef data3 As Object, _
                                <Out()> ByRef data4 As Object, <Out()> ByRef data5 As Object, <Out()> ByRef data6 As Object, <Out()> ByRef data7 As Object, _
                                <Out()> ByRef data8 As Object, <Out()> ByRef data9 As Object, <Out()> ByRef data10 As Object, _
                                 <Out()> ByRef data11 As Object, <Out()> ByRef data12 As Object, <Out()> ByRef data13 As Object, <Out()> ByRef data14 As Object, _
                                  <Out()> ByRef data15 As Object, <Out()> ByRef data16 As Object, <Out()> ByRef data17 As Object, <Out()> ByRef data18 As Object, _
                                   <Out()> ByRef data19 As Object, <Out()> ByRef data20 As Object)
        Dim datarow As System.Data.DataRow, i As Integer, col As Integer
        datarow = DirectCast(Row, System.Data.DataRow)
        i = (datarow.Table.Columns.Count)
        If i >= 19 Then
            col = 19
        Else
            col = i
        End If
        For i = 0 To col - 1
            If i = 0 AndAlso Not IsDBNull(datarow.Item(i)) Then data1 = datarow.Item(i)
            If i = 1 AndAlso Not IsDBNull(datarow.Item(i)) Then data2 = datarow.Item(i)
            If i = 2 AndAlso Not IsDBNull(datarow.Item(i)) Then data3 = datarow.Item(i)
            If i = 3 AndAlso Not IsDBNull(datarow.Item(i)) Then data4 = datarow.Item(i)
            If i = 4 AndAlso Not IsDBNull(datarow.Item(i)) Then data5 = datarow.Item(i)
            If i = 5 AndAlso Not IsDBNull(datarow.Item(i)) Then data6 = datarow.Item(i)
            If i = 6 AndAlso Not IsDBNull(datarow.Item(i)) Then data7 = datarow.Item(i)
            If i = 7 AndAlso Not IsDBNull(datarow.Item(i)) Then data8 = datarow.Item(i)
            If i = 8 AndAlso Not IsDBNull(datarow.Item(i)) Then data9 = datarow.Item(i)
            If i = 9 AndAlso Not IsDBNull(datarow.Item(i)) Then data10 = datarow.Item(i)
            If i = 10 AndAlso Not IsDBNull(datarow.Item(i)) Then data11 = datarow.Item(i)
            If i = 11 AndAlso Not IsDBNull(datarow.Item(i)) Then data12 = datarow.Item(i)
            If i = 12 AndAlso Not IsDBNull(datarow.Item(i)) Then data13 = datarow.Item(i)
            If i = 13 AndAlso Not IsDBNull(datarow.Item(i)) Then data14 = datarow.Item(i)
            If i = 14 AndAlso Not IsDBNull(datarow.Item(i)) Then data15 = datarow.Item(i)
            If i = 15 AndAlso Not IsDBNull(datarow.Item(i)) Then data16 = datarow.Item(i)
            If i = 16 AndAlso Not IsDBNull(datarow.Item(i)) Then data17 = datarow.Item(i)
            If i = 17 AndAlso Not IsDBNull(datarow.Item(i)) Then data18 = datarow.Item(i)
            If i = 18 AndAlso Not IsDBNull(datarow.Item(i)) Then data19 = datarow.Item(i)
            If i = 19 AndAlso Not IsDBNull(datarow.Item(i)) Then data20 = datarow.Item(i)
        Next
    End Sub
    Private Shared Function getDataSource(ByVal _Connection As System.Data.SqlClient.SqlConnection, ByVal dataSource As String, ByVal MaxRows As Int32, ByVal Fields As String) As System.Data.DataSet
        Dim _dataSource As DataSet
        Dim dataCommand As System.Data.SqlClient.SqlCommand
        Dim datatable As System.Data.DataTable

        If _Connection.State <> ConnectionState.Open Then _Connection.Open()
        _dataSource = New System.Data.DataSet
        dataCommand = _Connection.CreateCommand()
        If dataCommand.Connection.State <> ConnectionState.Open Then dataCommand.Connection.Open()
        '如果传入的数据源为空的话,则以Select 1 as Data代替之.这样处理的原因时,用户可能只是想进行一个简单的四则运算,并不需要数据库中的其他数据.
        If Trim(dataSource) = "" Then dataSource = "Select 1 as Data"
        '只取主数据源的100行
        'get the top 100 rows only
        If MaxRows < 0 Then MaxRows = 100
        dataSource = System.Text.RegularExpressions.Regex.Replace(dataSource, "^\s*select\s*(top\s*\(?(\d|@\w)+\)?)?", "Select Top " & MaxRows & " ", RegexOptions.IgnoreCase)
        '取出主数据源
        dataCommand.CommandText = dataSource
        datatable = DataReadToDataTable(dataCommand.ExecuteReader())
        '如果记录集为空,则允许表中所有列为NULL
        If datatable.Rows.Count = 0 Then
            For i As Integer = 0 To datatable.Columns.Count - 1
                datatable.Columns(i).AllowDBNull = True
            Next
        End If
        datatable.TableName = "DataSource"

        datatable.ExtendedProperties("SQL") = dataSource
        '如果没有行,则添加一行
        'if there is no rows then add a row
        'If datatable.Rows.Count = 0 Then datatable.Rows.Add(System.DBNull.Value)
        _dataSource.Tables.Add(datatable)
        '取出字段映射,只读取50行
        'get the fields map,only read the first 50 rows
        Fields = System.Text.RegularExpressions.Regex.Replace(Fields, "^\s*select\s*(top\s*\(?(\d|@\w)+\)?)?", "Select Top 300 ", RegexOptions.IgnoreCase)
        If Fields <> "" Then
            dataCommand.CommandText = Fields
            datatable = DataReadToDataTable(dataCommand.ExecuteReader())
            datatable.TableName = "Fields"
            datatable.ExtendedProperties("SQL") = Fields
            _dataSource.Tables.Add(datatable)
        End If
        Return _dataSource
    End Function
    <Microsoft.SqlServer.Server.SqlProcedure()> _
    Public Shared Function ExecuteNoQuery(ByVal Language As SqlInt16, ByVal Expression As SqlString, _
                                         ByVal DataSource As SqlString, ByVal Filter As SqlString, ByVal MaxRows As SqlInt32, ByVal Fields As SqlString, _
                                                 ByVal AutoAddQuotes As SqlBoolean) As SqlTypes.SqlInt32
        ' 在此处添加您的代码
        Dim _Datasource As DataSet
        Dim Scripts As ScriptTree
        Dim dataTable As New System.Data.DataTable
        If Expression = "" Then Return 0
        '_Datasource = New System.Data.DataSet
        Using Connection = New System.Data.SqlClient.SqlConnection("context connection=true")
            '取得记录集
            _Datasource = getDataSource(Connection, DataSource.Value, MaxRows.Value, Fields.Value)
            '替换关键字
            Expression = New KeywordsProcessor(Expression.Value, _Datasource, Filter.Value, AutoAddQuotes.Value).Expression()
            '分割子表达式
            'Scripts = SplitExpression(Expression.Value)
            '创建脚本树
            Scripts = New ScriptTree(Expression.Value, _Datasource)

            Dim retobj As System.Data.DataRow
            '从脚本树取得值
            'get value from the script tree
            retobj = Scripts.getValue(Connection)
            If Not (retobj Is Nothing) Then
                Return CType(retobj.Item(0), SqlInt32)
            Else
                Return 0
            End If
            Scripts = Nothing
        End Using
    End Function
    Protected Overrides Sub Finalize()
        MyBase.Finalize()

    End Sub
End Class
