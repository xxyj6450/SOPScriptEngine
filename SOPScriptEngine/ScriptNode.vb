Imports System.Xml.XPath
'脚本节点
Public Class ScriptNode
    Private _Cotent As String
    Private _Languate As String
    Private _ScriptType As String
    Private _Connection As System.Data.IDbConnection 'System.Data.SqlClient.SqlConnection
    Private _ScriptText As String
    Private _NodeID As String
    Private _Path As String
    Private _ParentNode As ScriptNode
    Friend _ChildrenCount As Integer
    Private _PreOperator As String
    Friend _Children As ScriptTree
    Public Sub New(ByVal Value As String)
        _ScriptText = Value
    End Sub
    Public Sub New(ByVal Value As String, ByVal ParentNodeValue As ScriptNode)
        _ScriptText = Value
        _ParentNode = ParentNodeValue
    End Sub
    '计算值
    Public Function getValue(ByVal Connection As System.Data.IDbConnection) As Object
        Return ""
    End Function
    '连接
    Public Property Connection() As System.Data.IDbConnection
        Get
            Return _Connection
        End Get
        Set(ByVal value As System.Data.IDbConnection)
            _Connection = value
        End Set
    End Property
    '表达式
    Public Property ScriptText() As String
        Get
            Return _ScriptText
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    '节点编码
    Public ReadOnly Property NodeID() As String
        Get
            If _NodeID = "" Then _NodeID = Guid.NewGuid.ToString
            Return _NodeID
        End Get

    End Property
    '路径
    Public ReadOnly Property Path() As String
        Get
            Return _Path
        End Get
    End Property
    '父节点
    Public Property ParentNode() As ScriptNode
        Get
            ParentNode = _ParentNode
        End Get
        Set(ByVal value As ScriptNode)
            _ParentNode = value
            _Path = value.Path & "/" & _NodeID
        End Set

    End Property
    Public ReadOnly Property ChildrenCount() As Integer
        Get
            Return _ChildrenCount
        End Get
    End Property
    Public Property Children() As ScriptTree
        Get
            Return _Children
        End Get
        Set(ByVal value As ScriptTree)
            _Children = value
        End Set
    End Property

End Class
