Imports System.Data.SqlClient
Imports System.Data

#Region " SQL Server DataTypes "
Public Class DataTypeSQL
    Public Sub New()
    End Sub
End Class

Public Class DateTimeSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, DateTime)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, DateTime)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, DateTime)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class IntSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Int64)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class VarcharSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) OrElse cValue = Nothing Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.String)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) OrElse Value = Nothing Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.String)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) OrElse cValue = Nothing Then
                Return String.Empty
            Else
                Return CType(cValue, System.String)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class BitSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                If Value Then
                    cValue = 1
                Else
                    cValue = 0
                End If
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Int64)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class MoneySQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                'Return CType(cValue, System.Decimal)
                Return FormatNumber(CType(cValue, System.Decimal), 2)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Decimal)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                'Return CType(cValue, System.Decimal)
                Return FormatNumber(CType(cValue, System.Decimal), 2)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class

Public Class FloatSQL
    Inherits DataTypeSQL
    Private cValue As Object = DBNull.Value
    Public Property Value()
        Get
            If IsDBNull(cValue) Then
                Return DBNull.Value
            Else
                Return CType(cValue, System.Double)
            End If
        End Get
        Set(ByVal Value As Object)
            If IsDBNull(Value) Then
                cValue = DBNull.Value
            Else
                cValue = CType(Value, System.Double)
            End If
        End Set
    End Property
    Public ReadOnly Property ToText()
        Get
            If IsDBNull(cValue) Then
                Return String.Empty
            Else
                Return CType(cValue, System.Double)
            End If
        End Get
    End Property
    Public Sub New(ByVal InitValue As Object)
        Value = InitValue
    End Sub
End Class
#End Region

Public Class DBase
    Public Function GetConvertDTSql(ByVal Sql As String, ByVal ServerIPAddress As String, ByVal DBName As String) As DataTable
        Dim Common As New Common
        Dim DataAdapter As Data.SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        Dim SqlCmd As New Data.SqlClient.SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New Data.SqlClient.SqlConnection(GetConnectionString(ServerIPAddress, DBName))
        DataAdapter = New Data.SqlClient.SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return Common.GetDTSql(dt)
    End Function


    Public Function GetDT(ByVal Sql As String, ByVal ServerIPAddress As String, ByVal DBName As String) As DataTable
        Dim DataAdapter As Data.SqlClient.SqlDataAdapter
        Dim dt As New DataTable
        Dim SqlCmd As New Data.SqlClient.SqlCommand(Sql)
        SqlCmd.CommandType = CommandType.Text
        SqlCmd.Connection = New Data.SqlClient.SqlConnection(GetConnectionString(ServerIPAddress, DBName))
        DataAdapter = New Data.SqlClient.SqlDataAdapter(SqlCmd)
        DataAdapter.Fill(dt)
        DataAdapter.Dispose()
        SqlCmd.Dispose()
        Return dt
    End Function

    Public Sub ExecuteNonQueryAllDB(ByVal Sql As String, ByVal ServerIPAddress As String, ByVal DBName As String)
        Dim SqlConnection1 As New Data.SqlClient.SqlConnection(GetConnectionString(ServerIPAddress, DBName))
        SqlConnection1.Open()
        Dim SqlCmd As New System.Data.SqlClient.SqlCommand(Sql, SqlConnection1)
        SqlCmd.CommandType = System.Data.CommandType.Text
        SqlCmd.ExecuteNonQuery()
        SqlCmd.Dispose()
        SqlConnection1.Close()
    End Sub

    Public Function DoesTableExist(ByVal ServerIPAddress As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim sb As New System.Text.StringBuilder
        Dim dt As DataTable

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")

        dt = GetDT(sb.ToString, ServerIPAddress, DBName)
        If dt.Rows(0)(0) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DoesDatabaseAndTableExist(ByVal ServerIPAddress As String, ByVal DBName As String, ByVal TableName As String) As Boolean
        Dim Sql As String

        Sql = DoesTableExistSql(DBName, TableName)
        Dim CmdAsst As New CmdAsst(Me, CommandType.Text, Sql, ServerIPAddress, DBName)
        Dim QueryPack As CmdAsst.QueryPack
        QueryPack = CmdAsst.Execute
        If QueryPack.Success Then
            If QueryPack.dt.rows(0)(0) = 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function DoesTableExistSql(ByVal DBName As String, ByVal TableName As String) As String
        Dim sb As New System.Text.StringBuilder

        sb.Append("USE " & DBName & " ")
        sb.Append("IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='")
        sb.Append(TableName)
        sb.Append("') ")
        sb.Append("SELECT Results = 1 ")
        sb.Append("ELSE ")
        sb.Append("SELECT Results = 0")
        Return sb.ToString
    End Function

    Public Function GetFieldList(ByVal ServerIPAddress As String, ByVal DBName As String, ByVal TableName As String) As DataTable
        Dim dt As DataTable
        Dim Sql As New System.Text.StringBuilder

        'Sql.append("SELECT * from information_schema.tables where table_type='BASE TABLE'")

        Sql.Append("SELECT column_name,ordinal_position,column_default,data_type, ")
        Sql.Append("Is_nullable from information_schema.columns ")
        Sql.Append("WHERE table_name='" & TableName & "'")
        Return GetDT(Sql.ToString, ServerIPAddress, DBName)
    End Function

    Public Function GetDTWithQueryPack(ByVal Sql As String, ByVal ServerIPAddress As String, ByVal DBName As String) As CmdAsst.QueryPack
        Dim MyResults As New Results
        Dim CmdAsst As CmdAsst
        Dim QueryPack As CmdAsst.QueryPack
        CmdAsst = New CmdAsst(Me, CommandType.Text, Sql, ServerIPAddress, DBName)
        QueryPack = CmdAsst.Execute

        MyResults.Success = QueryPack.Success
        If Not QueryPack.Success Then
            MyResults.Msg = "Unable to update record."
        End If
        Return QueryPack
    End Function

    Public Function GetConnectionString(ByVal ServerIPAddress As String, ByVal DBName As String) As String
        Dim Template As String
        'Template = "user id=BVI_SQL_SERVER;password=noisivtifeneb;database=|;server=~"
        Template = "user id=sa;password=susquehanna;database=|;server=~"
        Template = Replace(Template, "|", DBName)
        Template = Replace(Template, "~", ServerIPAddress)
        Return Template
    End Function

End Class

Public Class CmdAsst
    Private cSqlCmd As SqlClient.SqlCommand
    Private cDBName As String
    Private cServerIPAddress As String
    Private cDBase As DBase

    Public Sub New(ByRef DBase As DBase, ByVal CmdType As CommandType, ByRef SPNameOrSql As String, ByVal ServerIPAddress As String, ByVal DBName As String)
        cSqlCmd = New SqlClient.SqlCommand(SPNameOrSql)
        cSqlCmd.CommandType = CmdType
        cDBName = DBName
        cServerIPAddress = ServerIPAddress
        cDBase = DBase
    End Sub

#Region " Add Parameters "
    Public Sub AddInt(ByVal VarName As String, ByVal Value As Object)
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.ParameterDirection.Input, False, CType(10, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
        cSqlCmd.Parameters("@" & VarName).Value = Value
    End Sub
    Public Sub AddBit(ByVal VarName As String, ByVal Value As Object)
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Bit, 1"))
        cSqlCmd.Parameters("@" & VarName).Value = Value
    End Sub
    Public Sub AddDateTime(ByVal VarName As String, ByVal Value As Object) ' datetime
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.DateTime, 8"))
        cSqlCmd.Parameters("@" & VarName).Value = Value
    End Sub
    Public Sub AddMoney(ByVal VarName As String, ByVal Value As Object) ' double
        'Dim DblValue As Double
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Money, 4, System.Data.ParameterDirection.Input, False, CType(19, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
        If Not IsDBNull(Value) Then
            'cSqlCmd.Parameters("@" & VarName).Value = Double.Parse(Value)
            'cSqlCmd.Parameters("@" & VarName).Value = CType(Value, System.Double)
            'DblValue = Value
            cSqlCmd.Parameters("@" & VarName).Value = Value
        Else
            cSqlCmd.Parameters("@" & VarName).Value = 0 'DBNull.Value
        End If
    End Sub
    Public Sub AddVarChar(ByVal VarName As String, ByVal Value As Object, ByVal Length As Integer)
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.VarChar, "))
        cSqlCmd.Parameters("@" & VarName).Value = Value
    End Sub
    Public Sub AddFloat(ByVal VarName As String, ByVal Value As Object) ' double
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.Float, 8, System.Data.ParameterDirection.Input, False, CType(15, Byte), CType(0, Byte), """", System.Data.DataRowVersion.Current, Nothing"))
        cSqlCmd.Parameters("@" & VarName).Value = Value
    End Sub

    Public Sub AddGUID(ByVal VarName As String, ByVal Value As Object)
        cSqlCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@" & VarName, "System.Data.SqlDbType.UniqueIdentifier, "))
        cSqlCmd.Parameters("@" & VarName).Value = New Guid(Value.ToString)
    End Sub
#End Region

    Public Function GetDT() As DataTable
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        cSqlCmd.Connection = New SqlConnection(cDBase.GetConnectionString(cServerIPAddress, cDBName))
        DataAdapter = New SqlDataAdapter(cSqlCmd)
        DataAdapter.Fill(dt)
        DataAdapter.Dispose()
        cSqlCmd.Dispose()
        Return dt
    End Function

    Public Function Execute() As QueryPack
        Dim QueryPack As New QueryPack
        Dim DataAdapter As SqlDataAdapter
        Dim dt As New DataTable

        cSqlCmd.Connection = New SqlConnection(cDBase.GetConnectionString(cServerIPAddress, cDBName))
        DataAdapter = New SqlDataAdapter(cSqlCmd)
        Try
            DataAdapter.Fill(dt)
            QueryPack.Success = True
            QueryPack.dt = dt
        Catch ex As Exception
            QueryPack.Success = False
            QueryPack.TechErrMsg = ex.Message
        End Try
        DataAdapter.Dispose()
        cSqlCmd.Dispose()
        Return QueryPack
    End Function

    Public Class QueryPack
        Private cReturnDataTable As Boolean
        Private cReturnDataSet As Boolean
        Private cSuccess As Boolean
        Private cGenErrMsg As String
        Private cTechErrMsg As String
        Private cdt As DataTable
        Private cds As DataSet

        Public Property Success()
            Get
                Return cSuccess
            End Get
            Set(ByVal Value)
                cSuccess = Value
            End Set
        End Property

        Public ReadOnly Property GenErrMsg()
            Get
                Return GenErrMsg
            End Get
        End Property
        Public Property TechErrMsg()
            Get
                Return cTechErrMsg
            End Get
            Set(ByVal Value)
                cTechErrMsg = Value
            End Set
        End Property
        Public Property dt()
            Get
                Return cdt
            End Get
            Set(ByVal Value)
                cdt = Value
            End Set
        End Property
        Public Property ds()
            Get
                Return cds
            End Get
            Set(ByVal Value)
                cds = Value
            End Set
        End Property
    End Class
End Class