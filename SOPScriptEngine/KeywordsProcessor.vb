Imports System.Text.RegularExpressions
Public Class KeywordsProcessor
    ' Private _Expression As String
    Private _DataSource As System.Data.DataSet
    Private _Script As String
    Private _Content As String
    Private _Filter As String
    Private _AutoAddQuotes As Boolean
    'Private _ScriptText As String
    Public Sub New(ByVal Script As String, ByVal Datasource As System.Data.DataSet, ByVal _ActiveRow As String, ByVal blnAutoAddQuotes As Boolean)
        _Script = Script
        _DataSource = Datasource
        _Filter = _ActiveRow
        _AutoAddQuotes = blnAutoAddQuotes
    End Sub
    '返回表达式结果.Process参数表示是否进行一次关键字处理.默认为False.
    '调用此函数时,若发现尚未处理过关键字,则自动处理一次.
    Public Function Expression(Optional ByVal Process As Boolean = False) As String
        If _Content = "" Or Process = True Then _Content = ReplaceKeywords(_Script)
        Return _Content
    End Function
    '是否自动添加单引号.
    Public Property AutoAddQuotes() As Boolean
        Get
            AutoAddQuotes = _AutoAddQuotes
        End Get
        Set(ByVal value As Boolean)
            _AutoAddQuotes = value
        End Set
    End Property
    '替换关键字,需要先将表达式中的特殊字段做名称替换和值的替换
    Private Function ReplaceKeywords(ByVal Expression As String) As String
        Dim myMatchEvaluator As System.Text.RegularExpressions.MatchEvaluator
        '先字段名
        'If _Datasource.Tables("Fields").Rows.Count <> 0 And System.Text.RegularExpressions.Regex.IsMatch(Expression, "#(\d:)?(\w+)#") = True Then
        'myMatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceValueDelegate)
        'Expression = System.Text.RegularExpressions.Regex.Replace(Expression, "((\d:)?[\u4e00-\u9fa5])+", myMatchEvaluator)
        'myMatchEvaluator = Nothing
        'End If
        '再替换值
        myMatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceValueDelegate)
        Expression = System.Text.RegularExpressions.Regex.Replace(Expression, "(^|\+|-|\b|=|,|;|\()?(\s*)(')?&(?:(\d):)?([\w\s]+)&(\2)?(?:\[(.+?)\])?(\s*)(\+|-|=|,|b|;|\)|$)?", _
                                                                  myMatchEvaluator, RegexOptions.IgnoreCase)
        myMatchEvaluator = Nothing
        '替换字段
        myMatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceFieldsDelegate)
        Expression = System.Text.RegularExpressions.Regex.Replace(Expression, "#(\d:)?([\w\s]+)#", myMatchEvaluator, RegexOptions.IgnoreCase)
        myMatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ComputeExpression)
        Expression = System.Text.RegularExpressions.Regex.Replace(Expression, "(?:([\w|.]+?)::)?(\w+)\s*{(.+?)}(?:\[(.+?)\])?", _
                                                                  myMatchEvaluator, RegexOptions.IgnoreCase)
        myMatchEvaluator = New System.Text.RegularExpressions.MatchEvaluator(AddressOf ReplaceFilter)
        Expression = System.Text.RegularExpressions.Regex.Replace(Expression, "(^|\+|-|=|,|;|\b|\()\s*((?:sum|avg|count|min|max|var|stdev)\((?:[\w\s\+\-\*\/]+?)\))(?:\[(.+?)\])\s*(\+|-|=|,|;|\b|\)|$)", _
                                                          myMatchEvaluator, RegexOptions.IgnoreCase)
        Return Expression
    End Function
    '替换值,将&&括起来的串转换成字段的值
    '替换的过程是:
    '   0.取出字段名与表序号
    '   1.判断此字段是否存在数据表，若存在，直接输出字段名与表序号
    '   2.若不存在，则尝试从字段映射表中找字段，若字段映射表不存在，则抛出异常
    '       3.在字段映射表中找此字段
    '           3.1 先按字段序号+字段名查找 字段序号为0是要置空，将数据得到的数据插入数组
    '           3.2 若上一步无数据，则再按字段名查找，将得到的数据插入数组。本次虽然没有直接用字段序号，但是字段序号仍会用于定位数据表
    '           3.3 若上一步仍无数据，则用模糊匹配方式查找，字段映射表中要有设置模糊匹配才行，将查找结果插入数组，若未匹配到数据，则抛出异常。
    '       4.若第3步匹配到了多行，则说明字段映射表有重复，抛出异常。
    '       5.否则记录下此时的字段名与表序号。
    '   6.验证第1步或第5步返回的字段名与表序号是否在数据表中真的存在，匹配过程会分别用第0步的表序号与第5步的表序号来匹配数据表字段，优先使用第0步的表序号，若未匹配到，则抛出异常
    'groups(0):_Match本身
    'groups(1):前置运算符,可为:+-=,;(
    'groups(2):运算符前空白符
    'groups(3):单引号
    'groups(4):前置字段标志,以数字加冒号表示
    'groups(5):字段内容
    'groups(6):单引号
    'groups(7):[过滤条件]
    'groups(8):运算符后空白符
    'groups(9):后置运算符  
    Private Function ReplaceValueDelegate(ByVal _Match As System.Text.RegularExpressions.Match) As String
        Dim num As String, _Field As String, _FieldValue As String, _FieldName As String
        Dim groups As System.Text.RegularExpressions.GroupCollection = _Match.Groups
        Dim dataRows() As System.Data.DataRow
        Dim _Filter1 As String, m_Filter As String
        num = groups(4).Value.ToString
        _Field = groups(5).Value
        _Filter1 = groups(7).Value
        '还原字段名
        _FieldName = ReplaceFields(_Field, num, _Match.Value)
        '如果活动行不为空
        If _Filter <> "" Then
            m_Filter = _Filter
            '如果活动行不为空,则过滤不为空,则将活动行与过滤条件连接起来.如果过滤为空,则不处理
            If _Filter1 <> "" Then
                m_Filter = m_Filter & " AND " & _Filter1
            End If
        Else
            '如果活动行为空,而过滤条件不为空,则m_filter为过滤条件.如果过滤条件也为空,那不做处理,m_Filter为空
            If _Filter1 <> "" Then
                m_Filter = _Filter1
            End If
        End If
        dataRows = _DataSource.Tables("DataSource").Select(m_Filter)
        '如果未匹配到,则按原样返回
        If dataRows.Length = 0 Then
            Return groups(1).Value & groups(2).Value & groups(3).Value & "NULL" & groups(6).Value & groups(8).Value & groups(9).Value '_Match.Value
        Else
            '否则替换值,只用首行判断
            _FieldValue = dataRows(0)(_FieldName).ToString
        End If


        '如果数据不是数字,而且没有用单引号括起来(,并且在运算符前后,则补上单引号 'And groups(1).Value <> "" _)
        If Not IsNumeric(_FieldValue) And _AutoAddQuotes = True _
        And (groups(3).Value <> "'" And groups(6).Value <> "'") Then _FieldValue = "'" & _FieldValue & "'"
        Return groups(1).Value & groups(2).Value & groups(3).Value & _FieldValue & groups(6).Value & groups(8).Value & groups(9).Value
    End Function
    '替换字段
    Private Function ReplaceFieldsDelegate(ByVal _Match As System.Text.RegularExpressions.Match) As String
        Dim num As String, _Field As String, _FieldName As String
        Dim groups As System.Text.RegularExpressions.GroupCollection = _Match.Groups
        Dim dataRows() As System.Data.DataRow

        num = groups(1).Value.ToString
        _Field = groups(2).Value
        '如果表达式中的列不存时，则尝试从字段映射表中找映射字段
        _FieldName = ReplaceFields(_Field, num, _Match.Value)
        Return _FieldName
    End Function
    '在数据源中查找字段，并返回真实的字段名
    'SourceString:源字串，正则匹配到的数据，仅用于抛出错误，通知用户这个匹配失败。此参数不用于匹配字段。
    '_Field：用于匹配的字段名
    'num:用于匹配的表号
    '返回值：返回数据表中的真实字段。
    Private Function ReplaceFields(ByVal _Field As String, Optional ByVal num As String = "", Optional ByVal SourceString As String = "") As String
        Dim num1 As String, _FieldName As String, _Condition As String
        Dim dataRows() As System.Data.DataRow
        If Not _DataSource.Tables("DataSource").Columns.Contains(_Field & IIf(num = "0", "", num.ToString).ToString) Then
            '如果有字段映射表可匹配,则进行匹配
            If _DataSource.Tables.Contains("Fields") Then
                If num = "" Then
                    _Condition = ""
                Else
                    _Condition = " And TableNum=" & num
                End If
                '先按字段名+序号匹配,并且按表序号字段排序
                dataRows = _DataSource.Tables("Fields").Select("  FieldName ='" & _Field & "' And TableNum=" & IIf(num = "", "0", num).ToString, "TableNum ASC")

                If dataRows.Length = 0 Then
                    '若没有匹配到，再按字段名匹配
                    dataRows = _DataSource.Tables("Fields").Select("  FieldName ='" & _Field & "'" & _Condition, "TableNum ASC")
                    If dataRows.Length = 0 Then
                        '若没有匹配到，则再进行模糊匹配。用别名匹配中模糊设置来匹配。
                        dataRows = _DataSource.Tables("Fields").Select("'" & _Field & "' like FieldName" & _Condition, "TableNum ASC")
                        If dataRows.Length = 0 Then
                            Throw New ArgumentException("参数列" & _Field & "在字段映射表不存在。")
                        End If
                    End If
                End If
                '如果匹配到了多个列，则抛出异常。
                If dataRows.Length > 1 Then
                    Throw New ArgumentException("参数列" & _Field & "在字段映射表不唯一。")
                Else
                    '否则取得字段名和序号
                    _FieldName = dataRows(0)("fieldid").ToString
                    num1 = dataRows(0)("TableNum").ToString
                    num1 = IIf(num1 = "0", "", num1).ToString
                End If
            Else
                '如果没有字段映射表，则直接返回
                Throw New ArgumentException("参数列" & _Field & "不存在,字段映射表也不存在。")
            End If
        Else
            '若匹配到字段，则直接取出
            _FieldName = _Field
            num1 = IIf(num = "0", "", num).ToString
        End If
        If _DataSource.Tables("DataSource").Columns.Contains(_FieldName & num1.ToString) Then
            _FieldName = _FieldName & num1.ToString
        ElseIf _DataSource.Tables("DataSource").Columns.Contains(_FieldName & num.ToString) Then
            _FieldName = _FieldName & num.ToString
        Else
            '如果没有字段映射表，则直接返回
            Throw New ArgumentException("参数列" & _Field & "在数据表中不存在")
        End If
        Return _FieldName
    End Function
    '实现dataTable的Compute方法
    Private Function ComputeExpression(ByVal _match As System.Text.RegularExpressions.Match) As String
        Dim fun As String, Filter As String, Expression As String, ret As Object, ClassLib As String
        ClassLib = _match.Groups(1).Value
        fun = _match.Groups(2).Value
        Expression = _match.Groups(3).Value
        Filter = _match.Groups(4).Value
        If Trim(Filter) = "" Then Filter = "1=1"
        Select Case ClassLib
            Case "", "System.Data.DataTable"
                Select Case LCase(fun)
                    Case "compute"
                        ret = _DataSource.Tables("DataSource").Compute(Expression, Filter)
                    Case Else
                        Throw New ArgumentException(fun & "是未注册的函数。")
                End Select
            Case Else

                Throw New ArgumentException(ClassLib & "是未注册的类库。")
        End Select
        Return ret.ToString
    End Function
    '实现聚合函数的过滤
    Private Function ReplaceFilter(ByVal _match As System.Text.RegularExpressions.Match) As String
        Dim Filter As String, Expression As String, Ret As Object
        Expression = _match.Groups(2).Value
        Filter = _match.Groups(3).Value
        If Trim(Filter) = "" Then Filter = "1=1"
        Ret = _DataSource.Tables("DataSource").Compute(Expression, Filter)
        Return _match.Groups(1).Value + Ret.ToString + _match.Groups(4).Value
    End Function
    '解析子表达式,将表达式分析,得到子表达式
    Private Function SplitExpression(ByVal Expression As String) As System.Collections.Generic.SortedList(Of ScriptNode, Integer)
        Dim _Matches As System.Text.RegularExpressions.MatchCollection
        Dim _Match As System.Text.RegularExpressions.Match
        Dim Scripts As New System.Collections.Generic.SortedList(Of ScriptNode, Integer)
        Dim scriptNode As ScriptNode
        Dim i As Integer
        '判断是否有多个语句
        'If System.Text.RegularExpressions.Regex.IsMatch(Expression, "{.}") Then
        '    _Matches = System.Text.RegularExpressions.Regex.Matches(Expression, "")
        '    For Each _Match In _Matches
        '        scriptNode = New ScriptNode(Expression)
        '        i += 1
        '        Scripts.Add(scriptNode, i)
        '    Next
        'Else
        '    scriptNode = New ScriptNode(Expression)
        '    Scripts.Add(scriptNode, 1)
        'End If
        Scripts.Add(New ScriptNode(Expression), 0)
        Return Scripts
    End Function
End Class
