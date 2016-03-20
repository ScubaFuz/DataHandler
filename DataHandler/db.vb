Public Class db
    Private _DataProvider As String = "SQL"
    Private _DataLocation As String = Environment.MachineName
    Private _DatabaseName As String = "DataBase"
    Private _DataTableName As String = "Table"
    Private _LoginMethod As String = "WINDOWS"
    Private _LoginName As String = "sa"
    Private _Password As String = ""
    Private _ConnectionTimeout As Integer = 30
    Private _DataConnectionString As String
    Private _SqlConnection As New System.Data.SqlClient.SqlConnection
    Private _AccessConnection As New System.Data.OleDb.OleDbConnection
    Private _DataBaseChanged As Boolean = False
    Private _DatabaseOnline As Boolean = False
    Private _SqlVersion As Integer = 0

    Public dbMessage As String

#Region "Properties"
    Public Property DataProvider() As String
        Get
            Return _DataProvider
        End Get
        Set(ByVal Value As String)
            _DataProvider = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property DataLocation() As String
        Get
            Return _DataLocation
        End Get
        Set(ByVal Value As String)
            _DataLocation = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            Return _DatabaseName
        End Get
        Set(ByVal Value As String)
            _DatabaseName = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property DataTableName() As String
        Get
            Return _DataTableName
        End Get
        Set(ByVal Value As String)
            _DataTableName = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property LoginMethod() As String
        Get
            Return _LoginMethod
        End Get
        Set(ByVal Value As String)
            _LoginMethod = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property LoginName() As String
        Get
            Return _LoginName
        End Get
        Set(ByVal Value As String)
            _LoginName = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal Value As String)
            _Password = Value
            SetDataConnectionString()
        End Set
    End Property

    Public Property ConnectionTimeout() As Integer
        Get
            Return _ConnectionTimeout
        End Get
        Set(ByVal Value As Integer)
            _ConnectionTimeout = Value
            SetDataConnectionString()
        End Set
    End Property

    Public ReadOnly Property DataConnectionString() As String
        Get
            Return _DataConnectionString
        End Get
    End Property

    Public ReadOnly Property SqlConnection() As System.Data.SqlClient.SqlConnection
        Get
            dbMessage = Nothing
            'If String.IsNullOrEmpty(_SqlConnection.ConnectionString) Then
            '_SqlConnection.Dispose()
            '_SqlConnection = New System.Data.SqlClient.SqlConnection
            _SqlConnection.ConnectionString = _DataConnectionString
            'End If

            Return _SqlConnection
        End Get
    End Property

    Public ReadOnly Property AccessConnection() As System.Data.OleDb.OleDbConnection
        Get
            dbMessage = Nothing
            _AccessConnection.ConnectionString = _DataConnectionString
            Return _AccessConnection
        End Get
    End Property

    Public Property DataBaseChanged() As Boolean
        Get
            Return _DataBaseChanged
        End Get
        Set(ByVal Value As Boolean)
            _DataBaseChanged = Value
        End Set
    End Property

    Public Property DataBaseOnline() As Boolean
        Get
            Return _DatabaseOnline
        End Get
        Set(ByVal Value As Boolean)
            _DatabaseOnline = Value
        End Set
    End Property

    Public Property SqlVersion() As Integer
        Get
            Return _SqlVersion
        End Get
        Set(ByVal Value As Integer)
            _SqlVersion = Value
        End Set
    End Property

#End Region

#Region "Connection"
    Private Sub SetDataConnectionString()
        If SqlConnection.State = ConnectionState.Open Then SqlConnection.Close()
        If _DataProvider = Nothing Or _DataLocation = Nothing Or _DatabaseName = Nothing Or _LoginMethod = Nothing Then
            _DataConnectionString = Nothing
            Exit Sub
        End If
        If UCase(_LoginMethod) = "SQL" And (_LoginName = Nothing Or _Password = Nothing) Then
            _DataConnectionString = Nothing
            Exit Sub
        End If

        If UCase(_DataProvider) = "SQL" Then
            If UCase(_LoginMethod) = "SQL" Then
                _DataConnectionString = _
                "user id=" & _LoginName & ";" & _
                "data source=" & _DataLocation & ";" & _
                "persist security info=True;" & _
                "initial catalog=" & _DatabaseName & ";" & _
                "Connection Timeout=" & _ConnectionTimeout & ";" & _
                "password=""" & _Password & """"
            Else
                _DataConnectionString = _
                "integrated security=SSPI;" & _
                "data source=""" & _DataLocation & """;" & _
                "persist security info=False;" & _
                "Connection Timeout=" & _ConnectionTimeout & ";" & _
                "initial catalog=" & _DatabaseName & ""
            End If
        ElseIf UCase(_DataProvider) = "ACCESS" Then
            If UCase(_LoginMethod) = "WINDOWS" Then
                _LoginName = "Admin"
                _Password = ""
            End If
            _DataConnectionString = _
            "Jet OLEDB:Database Password=" & _Password & ";" _
            & "Data Source=""" & _DataLocation & "\" & _DatabaseName & ".mdb"";" _
            & "Password=" & _Password & ";" _
            & "Provider=""Microsoft.Jet.OLEDB.4.0"";" _
            & "Mode=ReadWrite;" _
            & "User ID=" & _LoginName & ";"
        End If
    End Sub

    Public Sub CheckDB()
        If UCase(_DataProvider) = "SQL" Then
            ConnectionTimeout = 5
            TestSQLConnection(SqlConnection)
            ConnectionTimeout = 30
        ElseIf UCase(_DataProvider) = "ACCESS" Then

        End If
    End Sub

    Public Sub TestSQLConnection(ByVal DataBase As System.Data.SqlClient.SqlConnection)
        Dim blnConnection As Boolean = False
        Try
            If DataBase.State = ConnectionState.Open Then DataBase.Close()
            If DataBase.State = ConnectionState.Closed Then DataBase.Open()
            If DataBase.State = ConnectionState.Open Then
                DataBaseOnline = True
            Else
                DataBaseOnline = False
            End If
            If DataBase.State = ConnectionState.Open Then DataBase.Close()
        Catch ex As Exception
            DataBaseOnline = False
        Finally
            If DataBase.State = ConnectionState.Open Then DataBase.Close()
        End Try
    End Sub

    Public Function GetSqlVersion() As Integer
        Dim strQuery As String = "exec [master].[dbo].[sp_server_info] 500"
        Dim dtsData As DataSet = QueryDatabase(strQuery, True)

        SqlVersion = 0
        If DataBaseOnline = False Then
            Return SqlVersion
        End If

        If dtsData Is Nothing Then Return 0
        If dtsData.Tables.Count = 0 Then Return 0
        If dtsData.Tables(0).Rows.Count = 0 Then Return 0

        Try
            For intRowCount As Integer = 0 To dtsData.Tables(0).Rows.Count - 1
                Dim strVersion As String = dtsData.Tables(0).Rows(intRowCount).Item("attribute_value")
                SqlVersion = strVersion.Substring(0, strVersion.IndexOf("."))
            Next
        Catch ex As Exception
            SqlVersion = 0
        End Try

        Return SqlVersion
    End Function
#End Region

#Region "Data Actions"
    Public Function QueryDatabase(ByVal SqlQuery As String, ByVal ReturnData As Boolean) As DataSet
        Dim objDataTemp As New DataSet

        If UCase(_DataProvider) = "SQL" Then
            If ReturnData = True Then
                objDataTemp = GetSqlData(SqlQuery, SqlConnection)
            ElseIf ReturnData = False Then
                UpdateSqlData(SqlQuery, SqlConnection)
            End If
        ElseIf UCase(_DataProvider) = "ACCESS" Then
            If ReturnData = True Then
                objDataTemp = GetAccessData(SqlQuery, AccessConnection)
            ElseIf ReturnData = False Then
                UpdateAccessData(SqlQuery, AccessConnection)
            End If
        End If
        QueryDatabase = objDataTemp
    End Function

    Private Function GetSqlData(ByVal mySelectQuery As String, ByVal DataBase As System.Data.SqlClient.SqlConnection) As DataSet
        dbMessage = Nothing
        Dim myCommand As New System.Data.SqlClient.SqlCommand(mySelectQuery, DataBase)
        Dim dataSet1 As New DataSet("DataSet1")
        If DataBase.State = ConnectionState.Closed Then DataBase.Open()
        Dim myReader As System.Data.SqlClient.SqlDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

        Dim dteInput As New DataTable("Table1")
        Dim RecordCount As Integer = 0
        'Dim Records As String

        If myReader.HasRows() = True Then
            While myReader.Read()
                Dim tRow As DataRow = dteInput.NewRow
                dteInput.Rows.Add(tRow)

                Dim i As Integer = 0
                If RecordCount = 0 Then
                    For i = 0 To myReader.FieldCount - 1
                        Dim Column As New DataColumn(myReader.GetName(i))
                        Column.DataType = myReader.GetFieldType(i)
                        dteInput.Columns.Add(Column)

                        Try
                            dteInput.Rows(RecordCount).Item(i) = myReader(i)
                        Catch ex As Exception
                            dbMessage = ex.Message
                        End Try
                    Next
                Else
                    For i = 0 To myReader.FieldCount - 1
                        Try
                            dteInput.Rows(RecordCount).Item(i) = myReader(i)
                        Catch ex As Exception
                            dbMessage = ex.Message
                        End Try
                    Next
                End If
                RecordCount += 1
            End While
        End If
        myReader.Close()
        Try
            If DataBase.State = ConnectionState.Open Then DataBase.Close()
        Catch ex As Exception
        End Try
        myCommand.Dispose()
        dataSet1.Tables.Add(dteInput)
        GetSqlData = dataSet1

    End Function

    Private Function GetAccessData(ByVal mySelectQuery As String, ByVal DataBase As System.Data.OleDb.OleDbConnection) As DataSet
        dbMessage = Nothing
        Dim myCommand As New System.Data.OleDb.OleDbCommand(mySelectQuery, DataBase)
        Dim dataSet1 As New DataSet("DataSet1")
        If DataBase.State = ConnectionState.Closed Then DataBase.Open()
        Dim myReader As System.Data.OleDb.OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

        Dim dteInput As New DataTable("Table1")
        Dim RecordCount As Integer = 0
        'Dim Records As String

        If myReader.HasRows() = True Then
            While myReader.Read()
                Dim tRow As DataRow = dteInput.NewRow
                dteInput.Rows.Add(tRow)

                Dim i As Integer = 0
                If RecordCount = 0 Then
                    For i = 0 To myReader.FieldCount - 1
                        Dim Column As New DataColumn(myReader.GetName(i))
                        Column.DataType = myReader.GetFieldType(i)
                        dteInput.Columns.Add(Column)

                        Try
                            dteInput.Rows(RecordCount).Item(i) = myReader(i)
                        Catch ex As Exception
                            dbMessage = ex.Message
                        End Try
                    Next
                Else
                    For i = 0 To myReader.FieldCount - 1
                        Try
                            dteInput.Rows(RecordCount).Item(i) = myReader(i)
                        Catch ex As Exception
                            dbMessage = ex.Message
                        End Try
                    Next
                End If
                RecordCount += 1
            End While
        End If
        myReader.Close()
        If DataBase.State = ConnectionState.Open Then DataBase.Close()
        myCommand.Dispose()
        dataSet1.Tables.Add(dteInput)
        GetAccessData = dataSet1

    End Function

    Private Sub UpdateSqlData(ByVal mySelectQuery As String, ByVal DataBase As System.Data.SqlClient.SqlConnection)
        dbMessage = Nothing
        Dim myCommand As New System.Data.SqlClient.SqlCommand(mySelectQuery, DataBase)
        If DataBase.State = ConnectionState.Closed Then DataBase.Open()
        Try
            Dim myReader As System.Data.SqlClient.SqlDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
            myReader.Close()
        Catch ex As Exception
            dbMessage = ex.Message
        End Try

        If DataBase.State = ConnectionState.Open Then DataBase.Close()
        myCommand.Dispose()
    End Sub

    Private Sub UpdateAccessData(ByVal mySelectQuery As String, ByVal DataBase As System.Data.OleDb.OleDbConnection)
        dbMessage = Nothing
        Dim myCommand As New System.Data.OleDb.OleDbCommand(mySelectQuery, DataBase)
        If DataBase.State = ConnectionState.Closed Then DataBase.Open()
        Try
            Dim myReader As System.Data.OleDb.OleDbDataReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
            myReader.Close()
        Catch ex As Exception
            dbMessage = ex.Message
        End Try

        If DataBase.State = ConnectionState.Open Then DataBase.Close()
        myCommand.Dispose()
    End Sub

    Public Function UploadSqlDataSet(ByVal objDataSet As DataSet) As Integer
        Dim intRecordsAffected As Integer = 0
        For Each objDataTable As DataTable In objDataSet.Tables
            intRecordsAffected += UploadSqlData(objDataTable)
        Next
        Return intRecordsAffected
    End Function

    Public Function UploadSqlData(ByVal objDataTable As DataTable) As Integer
        Dim intRecordsAffected As Integer = 0
        Dim bcp As System.Data.SqlClient.SqlBulkCopy = New System.Data.SqlClient.SqlBulkCopy(SqlConnection)
        If SqlConnection.State = ConnectionState.Closed Then SqlConnection.Open()
        bcp.DestinationTableName = DataTableName
        Dim reader As DataTableReader = objDataTable.CreateDataReader()
        intRecordsAffected = objDataTable.Rows.Count
        Try
            bcp.WriteToServer(reader)
        Catch ex As Exception
            Return 0
        End Try
        'bcp.Close()
        Try
            If SqlConnection.State = ConnectionState.Open Then SqlConnection.Close()
        Catch ex As Exception
        End Try
        Return intRecordsAffected
    End Function

    Public Sub WriteLog(ByVal LogText As String, ByVal EntryLevel As Integer, ByVal LogLevel As Integer, Optional ByVal Sender As String = "")
        Dim booLogItem As Boolean = False
        If Sender = Nothing Then Sender = Environment.MachineName

        If LogLevel >= EntryLevel Then
            LogText = Replace(LogText, "'", "~")
            Dim strQuery As String = "INSERT INTO [dbo].[tbl_Logging] ([LogDate],[Logtext],[ClientPC])"
            strQuery &= "VALUES(GetDate()," & LogText & "," & Sender & ")	"
            'Dim strQuery As String = "exec usp_LoggingHandle 'Ins',NULL,'" & LogText & "','" & Sender & "'"
            QueryDatabase(strQuery, False)
        End If

    End Sub

    Public Function ReplaceNulls(dtsInput As DataSet) As DataSet
        Dim dtsOutput As New DataSet
        Dim row As DataRow
        Dim col As DataColumn

        For Each Table As DataTable In dtsInput.Tables
            Dim newTable As DataTable = Table.Clone
            For Each colOrg As DataColumn In newTable.Columns
                colOrg.DataType = System.Type.GetType("System.String")
            Next
            For Each rowOrg As DataRow In Table.Rows
                newTable.ImportRow(rowOrg)
            Next
            dtsOutput.Tables.Add(newTable)

            For Each row In newTable.Rows
                For Each col In newTable.Columns
                    If row.IsNull(col) Then
                        Select Case Type.GetTypeCode(col.DataType)
                            Case TypeCode.Int32
                                row.Item(col) = 0
                            Case TypeCode.String
                                row.Item(col) = ""
                            Case Else
                                row.Item(col) = ""
                        End Select
                    End If
                Next
            Next
        Next
        Return dtsOutput
    End Function

#End Region


End Class
