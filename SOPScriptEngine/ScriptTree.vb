Imports System.Text.RegularExpressions
Public Class ScriptTree
    Private _ScriptNodes As New System.Collections.Generic.SortedDictionary(Of String, ScriptNode)
    Private _ScriptText As String
    Private _Content As String
    Private _Language As enumScriptLanguage
    Private _ScriptType As enumScriptType
    Private _DataSource As System.Data.DataSet

    '构造函数
    '2用数据源创建
    Public Sub New(ByVal _DataSource As System.Data.DataSet)

        Me._DataSource = _DataSource
    End Sub
    '3用脚本和数据源创建
    Public Sub New(ByVal Script As String, ByVal _DataSource As System.Data.DataSet)

        Me.DataSource = _DataSource
        Me.ScriptText = Script
    End Sub
    '脚本节点集
    Public Property ScriptNodes() As System.Collections.Generic.SortedDictionary(Of String, ScriptNode)
        Get
            Return _ScriptNodes
        End Get
        Set(ByVal value As System.Collections.Generic.SortedDictionary(Of String, ScriptNode))
            _ScriptNodes = value
        End Set
    End Property
    '添加节点
    '1.用节点用父节点添加
    Public Function AddNode(ByRef _Node As ScriptNode, ByVal ParentNode As ScriptNode) As ScriptNode

        Return _Node
    End Function
    '2.用节点和父节点ID添加
    Public Function AddNode(ByRef _Node As ScriptNode, ByVal ParentNodeID As String) As ScriptNode

        Return _Node
    End Function
    '3用节点文本和父节点添加
    Public Function AddNode(ByVal ScriptText As String, ByVal ParentNode As ScriptNode) As ScriptNode
        Dim _Node As New ScriptNode(ScriptText)
        _ScriptNodes.Add(_Node.NodeID, _Node)
        Return _Node
    End Function
    '4.用节点文本和父节点ID添加
    Public Function AddNode(ByVal ScriptText As String, ByVal ParentNodeID As String) As ScriptNode
        Dim _Node As New ScriptNode(ScriptText)
        _ScriptNodes.Add(_Node.NodeID, _Node)
        Return _Node
    End Function
    '根据ID获取节点
    Public Function getNodeByID(ByVal NodeID As String) As ScriptNode
        If _ScriptNodes.ContainsKey(NodeID) Then
            Return _ScriptNodes.Item(NodeID)
        Else
            Return Nothing
        End If
    End Function
    '数据源
    Public Property DataSource() As System.Data.DataSet
        Get
            Return _DataSource
        End Get
        Set(ByVal value As System.Data.DataSet)
            _DataSource = value
        End Set
    End Property

    '脚本代码
    Private Property ScriptText() As String
        Get
            Return _ScriptText
        End Get
        Set(ByVal value As String)
            _ScriptText = value
            '如果表达式是以Select开头,则为Select子句,直接执行
            '如果表达式是Where开头,则为Where子句,则将Where删除,补到Dataset的过滤表达式后
            '如果表达式中没有Select,=,<,>,<>,!=,in,exists,则认为是Select子句
            '如果表达式中有=,<,>,<>,!=,in,exists,Like,则必须在每个这些运算符有When才算Select子句,否则算Where子句,只判断第一个.

            If Regex.IsMatch(value, "^\s*select", RegexOptions.IgnoreCase) Then
                '1.匹配Select开头的Select子句
                _Content = value
                _ScriptType = enumScriptType.SelectScript
                Exit Property

            ElseIf Regex.IsMatch(value, "^\s*Where", RegexOptions.IgnoreCase) Then
                '2.匹配Where开头的Where子句
                If System.Text.RegularExpressions.Regex.IsMatch(_DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString, "\bwhere\b", RegexOptions.IgnoreCase) Then
                    _Content = _DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString & " AND " & System.Text.RegularExpressions.Regex.Replace(value, "^\s*Where", " ", RegexOptions.IgnoreCase)                '需要将Where替换掉
                Else
                    _Content = _DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString & " " & value
                End If
                _ScriptType = enumScriptType.WhereScript
                Exit Property
            ElseIf Regex.IsMatch(value, "^\s*Exists", RegexOptions.IgnoreCase) Then
                '2.匹配Exists开头的Where子句
                If System.Text.RegularExpressions.Regex.IsMatch(_DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString, "\bwhere\b", RegexOptions.IgnoreCase) Then
                    _Content = _DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString & " AND " & System.Text.RegularExpressions.Regex.Replace(value, "^\s*Where", " ", RegexOptions.IgnoreCase)                '需要将Where替换掉
                Else
                    _Content = _DataSource.Tables("DataSource").ExtendedProperties("SQL").ToString & " Where " & value
                End If
                _ScriptType = enumScriptType.WhereScript
                Exit Property
            ElseIf Regex.IsMatch(value, "^\s*!") Or Regex.IsMatch(value, "^\s*Update\s+", RegexOptions.IgnoreCase) Or Regex.IsMatch(value, "^\s*Delete\s+", RegexOptions.IgnoreCase) _
            Or Regex.IsMatch(value, "^\s*Exec(ute)?\s+", RegexOptions.IgnoreCase) Or Regex.IsMatch(value, "^\s*Insert\s+", RegexOptions.IgnoreCase) Or Regex.IsMatch(value, "^\s*Drop\s+", RegexOptions.IgnoreCase) _
            Or Regex.IsMatch(value, "^\s*Alter\s+", RegexOptions.IgnoreCase) Or Regex.IsMatch(value, "^\s*Create\s+", RegexOptions.IgnoreCase) Then
                '3.匹配!开头的存储过程
                _Content = value.Replace("!", "") ' System.Text.RegularExpressions.Regex.Replace(value, "^\s*!", " ")
                '如果是用分号分隔,则替换分号为逗号
                _Content = _Content.Replace(";", ",")
                _ScriptType = enumScriptType.Procedure
                Exit Property
            ElseIf System.Text.RegularExpressions.Regex.IsMatch(value, "(=|!=|<|>|<>|\bin\b|\bexists\b)", RegexOptions.IgnoreCase) Then
                '4匹配包含
                Dim match As System.Text.RegularExpressions.Match
                '取得第一个满足条件的匹配
                match = System.Text.RegularExpressions.Regex.Matches(value, "=|>=|<=|!=|<|>|<>|\b(in|exists|like)\b", RegexOptions.IgnoreCase)(0)
                If System.Text.RegularExpressions.Regex.IsMatch(value.Substring(0, match.Length + match.Index), _
                                                                "\bIIF\s*((\S+\s*(=|>=|<=|!=|<|>|<>|\b(in|like)\b))|\bexists\b)", RegexOptions.IgnoreCase) Then
                    '如果第一个=,!=,<,>,<>,exists,in,like关键字之前有Case When语句,则该语句为Select子句,否则为Where子句.
                    _ScriptType = enumScriptType.SimpleSelectScript
                Else
                    _ScriptType = enumScriptType.SimpleWhereScript

                End If
                _Content = value
            Else
                '5不是以Select/Where/!开头,且不包含=,!=,<,>,<>,exists,in,like关键子的,作为Select子句
                _ScriptType = enumScriptType.SimpleSelectScript
                _Content = value
            End If
        End Set
    End Property
    Public Function getValue(ByVal connection As System.Data.SqlClient.SqlConnection, Optional ByVal MultiLine As Boolean = False) As System.Data.DataRow
        Dim command As System.Data.SqlClient.SqlCommand

        Dim i As Integer
        Dim dataTable As System.Data.DataTable
        Dim ColumnArr As String()
        Dim column As System.Data.DataColumn
        Dim datarows As System.Data.DataRow()
        Dim dataRow As System.Data.DataRow
        '如果没有文本,则输出NULL
        If _Content = "" Then Return Nothing
        If _ScriptType = enumScriptType.Procedure Or _ScriptType = enumScriptType.SelectScript Or _ScriptType = enumScriptType.WhereScript Then
            command = New System.Data.SqlClient.SqlCommand
            command.Connection = connection
            If command.Connection.State <> ConnectionState.Open Then command.Connection.Open()
        End If
        Select Case _ScriptType
            '如果是存储过程,则直接执行
            Case enumScriptType.Procedure
                Dim m_AffectedRows As Integer
                '将单引号替换成双单引号
                _Content = Replace(_Content, "'", "''")
                command.CommandText = "exec  sp_executesql N'" & _Content & "'"
                'command.CommandType = CommandType.StoredProcedure
                '如果数据源不包含行,且表达式不包含动态参数,则执行之,若有动态参数则不执行.
                m_AffectedRows = command.ExecuteNonQuery()
                dataTable = New System.Data.DataTable
                dataTable.Columns.Add("ScriptExpressionColumn", Type.GetType("System.Int32"))
                dataTable.Columns("ScriptExpressionColumn").ExtendedProperties("ExtendID") = 0
                dataRow = dataTable.NewRow()
                dataRow("ScriptExpressionColumn") = m_AffectedRows
                dataTable.Rows.Add(dataRow)
                Return dataRow
                Return Nothing
            Case enumScriptType.SelectScript
                '如果是Select脚本,则执行Select语句,会与数据库交互
                command.CommandText = _Content
                dataTable = DataReadToDataTable(command.ExecuteReader(CommandBehavior.SingleRow))
                For i = 0 To dataTable.Columns.Count - 1
                    dataTable.Columns(i).ExtendedProperties("ExtendID") = i
                Next
                Return dataTable.Rows(0)
            Case enumScriptType.SimpleWhereScript
                'Where脚本,该脚本将作为DataTable的Select子句,只支持比较简单的语法,不与数据库交互,详情请查看MSDN中DataColumn的Expression属性.复杂的SQL语法需要使用Select方式
                If _DataSource Is Nothing Then Return Nothing
                '若过滤条件不为空,则补齐过滤条件
                If _DataSource.ExtendedProperties("Filter").ToString <> "" Then
                    If _Content = "" Then
                        _Content = _DataSource.ExtendedProperties("Filter").ToString
                    Else
                        _Content = _Content & " AND " & _DataSource.ExtendedProperties("Filter").ToString
                    End If

                End If
                datarows = _DataSource.Tables("DataSource").Select(_Content)
                dataTable = New System.Data.DataTable
                dataTable.Columns.Add("ScriptExpressionColumn", Type.GetType("System.Int32"))
                dataTable.Columns("ScriptExpressionColumn").ExtendedProperties("ExtendID") = 0
                dataRow = dataTable.NewRow()
                If _DataSource.Tables("DataSource").Rows.Count = 0 Then
                    dataRow("ScriptExpressionColumn") = System.DBNull.Value
                Else
                    If datarows.Length = 0 Then
                        dataRow("ScriptExpressionColumn") = 0
                    Else
                        dataRow("ScriptExpressionColumn") = 1
                    End If
                End If
                dataTable.Rows.Add(dataRow)
                Return dataRow
            Case enumScriptType.WhereScript
                command.CommandText = _Content
                dataTable = New System.Data.DataTable
                dataTable.Columns.Add("ScriptExpressionColumn", Type.GetType("System.Int32"))
                dataTable.Columns("ScriptExpressionColumn").ExtendedProperties("ExtendID") = 1
                dataRow = dataTable.NewRow()
                '如果数据源中无数据,则直接返回NULL
                If _DataSource.Tables("DataSource").Rows.Count = 0 Then
                    dataRow("ScriptExpressionColumn") = System.DBNull.Value
                Else
                    If DataReadToDataTable(command.ExecuteReader(CommandBehavior.SingleRow)).Rows.Count = 0 Then
                        dataRow("ScriptExpressionColumn") = 0
                    Else
                        dataRow("ScriptExpressionColumn") = 1
                    End If
                End If
                dataTable.Rows.Add(dataRow)
                Return dataTable.Rows(0)
            Case enumScriptType.SimpleSelectScript
                '简单脚本,该脚本使用DataColumn的Expression方式取值,不与数据库交互.
                dataTable = New System.Data.DataTable
                dataRow = dataTable.NewRow()

                ColumnArr = Split(_Content, ";")
                '添加列到DataSource，再将新加列的值取出，放到dataTable中，以供返回
                For i = 0 To ColumnArr.Length - 1
                    column = New System.Data.DataColumn("ScriptExpressionColumn" & i)
                    With column
                        .Expression = ColumnArr(i)
                        .ExtendedProperties.Add("ExtendID", i)

                    End With
                    _DataSource.Tables("DataSource").Columns.Add(column)
                    '将Datasource的计算列添加到datatable时需要判断该列是不是DBNull，因为数据库中列类型不兼容DBNull类型，遇到这种情况需要特殊处理，这里转为System.String类型处理，值依然是NULL
                    If _DataSource.Tables("DataSource").Rows(0).Item("ScriptExpressionColumn" & i).GetType.FullName = "System.DBNull" Then
                        dataTable.Columns.Add("ScriptExpressionColumn" & i, System.Type.GetType("System.String"))
                        dataRow("ScriptExpressionColumn" & i) = System.DBNull.Value
                    Else
                        dataTable.Columns.Add("ScriptExpressionColumn" & i, Type.GetType(_DataSource.Tables("DataSource").Rows(0).Item("ScriptExpressionColumn" & i).GetType.FullName))

                        dataRow("ScriptExpressionColumn" & i) = _DataSource.Tables("DataSource").Rows(0).Item("ScriptExpressionColumn" & i)
                    End If


                Next
                dataTable.Rows.Add(dataRow)
                'dataTable = _DataSource.Tables("DataSource").Copy
                ''清空Datatable的所有关系
                'If Not (dataTable.Constraints Is Nothing) Then dataTable.Constraints.Clear()
                'If Not (dataTable.ChildRelations Is Nothing) Then dataTable.ChildRelations.Clear()
                'If Not (dataTable.ParentRelations Is Nothing) Then dataTable.ParentRelations.Clear()

                'i = 0
                'While Not dataTable.Columns(i).ExtendedProperties.Contains("ExtendID")
                '    If dataTable.Columns.CanRemove(dataTable.Columns(i)) = True Then
                '        dataTable.Columns.Remove(dataTable.Columns(i))
                '    Else
                '        i = i + 1
                '    End If
                'End While
                ''如果表内无行,则添加

                Return dataTable.Rows(0)
        End Select
    End Function
    Public Sub Execute(ByVal connection As System.Data.SqlClient.SqlConnection)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        DataSource = Nothing
        _ScriptNodes = Nothing
    End Sub
End Class
