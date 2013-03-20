Imports System
Imports System.Text.RegularExpressions
Module Example
    Function CapText(ByVal m As Match) As String
        ' Get the matched string.
        Dim x As String = m.ToString()
        ' If the first char is lower case...
        If Char.IsLower(x.Chars(0)) Then
            ' Capitalize it.
            Return Char.ToUpper(x.Chars(0)) & x.Substring(1, x.Length - 1)
        End If
        Return x
    End Function

    Sub Main()

        Dim connstr As String
        Dim ds As New DataSet
        Dim sql As String
        connstr = "data source=192.168.16.49,9090;uid=sa;pwd=10180323zlxhyc;initial catalog=urp"
        sql = "Select top 1 price,price,price, price as price1,price as price2 From Unicom_orders a "
        Dim dap As New System.Data.SqlClient.SqlDataAdapter(sql, connstr)
        dap.Fill(ds)
        Debug.Print(ds.Tables(0).Rows.Count)
    End Sub
End Module