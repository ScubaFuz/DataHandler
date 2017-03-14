Imports System.IO
Imports System
Imports System.Data
Imports System.Math
Imports System.Collections
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Text
Imports System.Environment.SpecialFolder
Imports System.Net.Mail
Imports System.Xml.XPath

Public Class txt

#Region "General"
    Private _ErrorLevel As Integer = 0
    Private _ErrorMessage As String = ""

    Public Property ErrorLevel() As Integer
        Get
            Return _ErrorLevel
        End Get
        Set(ByVal Value As Integer)
            _ErrorLevel = Value
        End Set
    End Property

    Public Property ErrorMessage() As String
        Get
            Return _ErrorMessage
        End Get
        Set(ByVal Value As String)
            _ErrorMessage = Value
        End Set
    End Property

    Public Function ReplaceFirst(ByVal text As String, ByVal search As String, ByVal replace As String) As String
        Dim pos As Integer = text.IndexOf(search)
        If pos >= 0 Then
            Return text.Substring(0, pos) + replace + text.Substring(pos + search.Length)
        End If
        Return text
    End Function
#End Region

#Region "LogFile"
    Private _LogLevel As Integer = 1
    Private _LogLocation As String = "C:\"
    Private _LogFileName As String = "General"
    Private _Retenion As String = "Month"
    Private _AutoDelete As Boolean = False

    Public Property LogLevel() As Integer
        Get
            Return _LogLevel
        End Get
        Set(ByVal Value As Integer)
            If Value >= 0 And Value <= 5 Then
                _LogLevel = Value
            End If
        End Set
    End Property

    Public Property LogLocation() As String
        Get
            Return _LogLocation
        End Get
        Set(ByVal Value As String)
            If Value.ToLower = "database" Or CheckDir(Value) = True Then
                _LogLocation = Value
            End If
        End Set
    End Property

    Public Property LogFileName() As String
        Get
            Return _LogFileName
        End Get
        Set(ByVal Value As String)
            If CheckFileName(Value) = True Then
                _LogFileName = Value
            End If
        End Set
    End Property

    Public Function WriteLog(ByVal LogText As String, ByVal EntryLevel As Integer, Optional ByVal Sender As String = Nothing) As Boolean
        _Errormessage = ""
        Dim booLogItem As Boolean = False
        If Sender = Nothing Then Sender = Environment.MachineName

        Try
            If _LogLevel >= EntryLevel Then
                Dim strDate As String
                'intDate = Today.Year & [Enum].Format(GetType(Integer), Today.Month, "00") & [Enum].Format(GetType(Integer), Today.Day, "00")
                strDate = Today.ToString("yyyyMMdd")
                Dim objWriter As StreamWriter = File.AppendText(PathConvert(_LogLocation) & "\" & strDate & "_" & _LogFileName)
                'objWriter.WriteLine(Format(GetType(Integer), Now.Hour, "00") & ":" & Format(GetType(Integer), Now.Minute, "00") & ":" & Format(GetType(Integer), Now.Second, "00") & vbTab & LogText)
                objWriter.WriteLine(Now.ToString("HH:mm:ss") & vbTab & LogText)
                objWriter.Close()
                objWriter = Nothing
            End If
        Catch ex As Exception
            _Errormessage = ex.Message
            Return False
        End Try
        Return True
    End Function

    Public Property Retenion() As String
        Get
            Return _Retenion
        End Get
        Set(ByVal Value As String)
            _Retenion = Value
        End Set
    End Property

    Public Property AutoDelete() As Boolean
        Get
            Return _AutoDelete
        End Get
        Set(ByVal Value As Boolean)
            _AutoDelete = Value
        End Set
    End Property

#End Region

#Region "Input-Output"
    Private _InputFile As String
    Private _OutputFile As String
    Private _ImportFile As String
    Private _ExportFile As String

    Public Property InputFile() As String
        Get
            Return _InputFile
        End Get
        Set(ByVal Value As String)
            If CheckFileName(Value) = True Then
                _InputFile = Value
            End If
        End Set
    End Property

    Public Property OutputFile() As String
        Get
            Return _OutputFile
        End Get
        Set(ByVal Value As String)
            If CheckFileName(Value) = True Then
                _OutputFile = Value
            End If
        End Set
    End Property

    Public Property ImportFile() As String
        Get
            Return _ImportFile
        End Get
        Set(ByVal Value As String)
            _ImportFile = Value
        End Set
    End Property

    Public Property ExportFile() As String
        Get
            Return _ExportFile
        End Get
        Set(ByVal Value As String)
            _ExportFile = Value
        End Set
    End Property

    Public Function CreateDir(ByVal NewDir As String) As Boolean
        If NewDir.Length < 2 Then Return False
        If Directory.Exists(PathConvert(NewDir)) = True Then
            CreateDir = True
        Else
            Try
                Directory.CreateDirectory(PathConvert(NewDir))
                CreateDir = True
            Catch ex As Exception
                CreateDir = False
            End Try
        End If
    End Function

    Public Function CheckDir(ByVal Dir As String, Optional blnCreateDir As Boolean = False) As Boolean
        If Dir.Length < 2 Then Return False
        Dim myIO As New DirectoryInfo(PathConvert(Dir))
        Dim blnDirExists As Boolean = myIO.Exists()

        If blnCreateDir = True And blnDirExists = False Then
            Try
                If Dir.Contains("\") Then
                    Dim ParentDir As String = Dir.Substring(0, Dir.LastIndexOf("\"))
                    CheckDir(ParentDir, blnCreateDir)
                End If
                If CreateDir(Dir) = True Then blnDirExists = True
            Catch ex As Exception
                blnDirExists = False
            End Try
        End If
        Return blnDirExists
        myIO = Nothing
    End Function

    Public Function CheckFile(ByVal FileName As String) As Boolean
        If FileName.Length < 2 Then Return False
        Dim myIO As New FileInfo(PathConvert(FileName))
        Return myIO.Exists()
        myIO = Nothing
    End Function

    Public Function CheckFileName(ByVal FileName As String) As Boolean
        If FileName.Length < 2 Then Return False
        Dim strAllowedCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_@~#."
        Dim intCharCount As Integer
        Dim booCharacterOk As Boolean = True

        For intCharCount = 1 To FileName.Length
            If InStr(strAllowedCharacters, UCase(Mid(FileName, intCharCount, 1))) = 0 Then booCharacterOk = False
        Next
        Return booCharacterOk
    End Function

    Public Function CreateFile(ByVal InputText As String, ByVal FileName As String) As Boolean
        If FileName Is Nothing Then Return False
        If FileName = String.Empty Then Return False
        If FileName.Length < 2 Then Return False

        Dim intCheckDir As Integer = 0, strCheckdir As String = Nothing
        intCheckDir = FileName.IndexOf("\")
        If intCheckDir > -1 Then
            strCheckdir = FileName.Substring(0, FileName.LastIndexOf("\"))
        Else
            strCheckdir = System.Reflection.Assembly.GetCallingAssembly.Location.Substring(0, System.Reflection.Assembly.GetCallingAssembly.Location.LastIndexOf("\"))
            FileName = strCheckdir & "\" & FileName
        End If

        If CheckDir(strCheckdir, True) Then
            Try
                If FileCreate(InputText, FileName) = False Then Return False
                Return True
            Catch ex As Exception
                _Errormessage = ex.Message
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function ReadFile(ByVal FileName As String) As StreamReader
        If FileName.Length < 2 Then Return Nothing
        Dim myFileIO As New StreamReader(PathConvert(FileName))
        Return myFileIO
    End Function

    Public Function WriteFile(ByVal InputText As String, ByVal FileName As String) As Boolean
        If FileName.Length < 2 Then Return False
        Dim strCheckDir As String
        'Dim strCheckFile As String

        strCheckDir = Left(FileName, InStrRev(FileName, "\") - 1)

        If strCheckDir.Length = 0 Or CheckDir(strCheckDir) Then
            Try
                If CheckFile(FileName) Then
                    FileAdd(InputText, FileName)
                Else
                    FileCreate(InputText, FileName)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If

    End Function

    Private Function FileCreate(ByVal InputText As String, ByVal FileName As String) As Boolean
        If FileName.Length < 2 Then Return False
        Try
            Dim objWriter As StreamWriter = File.CreateText(PathConvert(FileName))
            objWriter.WriteLine(InputText)
            objWriter.Close()
            objWriter = Nothing
            Return True
        Catch ex As Exception
            _Errormessage = ex.Message
            Return False
        End Try
    End Function

    Private Sub FileAdd(ByVal WriteLine As String, ByVal FileName As String)
        If FileName.Length < 2 Then Exit Sub
        Dim objWriter As StreamWriter = File.AppendText(PathConvert(FileName))
        objWriter.WriteLine(WriteLine)
        objWriter.Close()
        objWriter = Nothing
    End Sub

    Public Function PathConvert(strPath As String) As String
        If strPath.Contains("%Documents%") Then
            strPath = strPath.Replace("%Documents%", Environment.GetFolderPath(MyDocuments))
        End If
        If strPath.Contains("%ProgramFiles%") Then
            strPath = strPath.Replace("%ProgramFiles%", Environment.GetFolderPath(ProgramFiles))
        End If
        If strPath.Contains("%ApplicationData%") Then
            strPath = strPath.Replace("%ApplicationData%", Environment.GetFolderPath(ApplicationData))
        End If
        If strPath.Contains("%CommonApplicationData%") Then
            strPath = strPath.Replace("%CommonApplicationData%", Environment.GetFolderPath(CommonApplicationData))
        End If
        If strPath.Contains("%Desktop%") Then
            strPath = strPath.Replace("%Desktop%", Environment.GetFolderPath(Desktop))
        End If
        Return strPath
    End Function

    Public Function DatasetCheck(dtsInput As DataSet, Optional intTable As Integer = 0) As Boolean
        Dim blnOK As Boolean = True

        Try
            If dtsInput Is Nothing Then Return False
            If dtsInput.Tables.Count = 0 Then Return False
            If dtsInput.Tables.Count < intTable + 1 Then Return False
            If dtsInput.Tables(intTable).Rows.Count = 0 Then Return False
        Catch ex As Exception
            Return False
        End Try

        Return blnOK
    End Function

    Public Function CsvToDataSet(strFileName As String, blnHasHeaders As Boolean, Optional Delimiter As String = ",", Optional QuoteValues As Boolean = False, Optional HeadersOnly As Boolean = False, Optional TextEncoding As String = "UTF8") As DataSet
        Dim dtsOutput As New DataSet
        Dim dttOutput As New DataTable
        dtsOutput.Tables.Add(dttOutput)

        Dim intRowCount As Integer = 0
        Dim intMaxColCount As Integer = 0
        ErrorLevel = 0
        ErrorMessage = ""

        Dim encInput As Text.Encoding = Text.Encoding.UTF8
        Select Case TextEncoding.ToUpper
            Case "UTF8"
                encInput = Encoding.UTF8
            Case "UTF7"
                encInput = Encoding.UTF7
            Case "UTF32"
                encInput = Encoding.UTF32
            Case "ASCII"
                encInput = Encoding.ASCII
            Case "UNICODE"
                encInput = Encoding.Unicode
            Case "BIGENDIANUNICODE"
                encInput = Encoding.BigEndianUnicode
        End Select
        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(strFileName, encInput)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.HasFieldsEnclosedInQuotes = QuoteValues
            MyReader.SetDelimiters(Delimiter)

            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    If intRowCount = 0 And blnHasHeaders = False Then
                        intMaxColCount = currentRow.Count
                        For intColumns As Integer = 1 To currentRow.Count
                            dttOutput.Columns.Add("col" & intColumns)
                        Next
                    End If
                    If intRowCount > 0 Or blnHasHeaders = False Then dttOutput.Rows.Add()

                    Dim currentField As String
                    Dim intColCount As Integer = 0
                    For Each currentField In currentRow
                        If intRowCount = 0 And blnHasHeaders = True Then
                            intMaxColCount = currentRow.Count
                            'Create Columns
                            dttOutput.Columns.Add(currentField)
                        ElseIf intRowCount > 0 And HeadersOnly = True Then
                            Exit While
                        Else
                            'fill datarow
                            If intColCount < intMaxColCount Then
                                dttOutput.Rows(dttOutput.Rows.Count - 1)(intColCount) = currentField
                            Else
                                ErrorLevel = -1
                                ErrorMessage = "To many columns For row " & intRowCount + 1 & ". Data may have been lost."
                                Console.WriteLine("Error: To many columns For row " & intRowCount + 1 & ". Data may have been lost.")
                                WriteLog("To many columns For row " & intRowCount + 1 & ". Data may have been lost.", 1)
                            End If
                        End If
                        intColCount += 1
                        'MsgBox(currentField)
                    Next
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    Console.WriteLine("Error: Line " & intRowCount + 1 & " Is Not valid And will be skipped. " & ex.Message)
                    WriteLog("Line " & intRowCount + 1 & " Is Not valid And will be skipped. " & ex.Message, 1)
                Catch ex As Exception
                    Console.WriteLine("Error: Line " & intRowCount + 1 & " Is Not valid And will be skipped. " & ex.Message)
                    WriteLog("Line " & intRowCount + 1 & " Is Not valid And will be skipped. " & ex.Message, 1)
                End Try
                intRowCount += 1
            End While
        End Using
        Return dtsOutput
    End Function

    Public Function DataSetToCsv(dttSource As DataTable, strFileName As String, Optional blnHasHeaders As Boolean = True, Optional Delimiter As String = ",", Optional QuoteValues As Boolean = False) As Boolean
        Try
            Using writer As IO.StreamWriter = New IO.StreamWriter(strFileName)
                If (blnHasHeaders) Then
                    Dim headerValues As System.Collections.Generic.IEnumerable(Of String) = dttSource.Columns.OfType(Of DataColumn).Select(Function(column) QuoteValue(column.ColumnName, QuoteValues))
                    writer.WriteLine(String.Join(Delimiter, headerValues))
                End If

                Dim items As System.Collections.Generic.IEnumerable(Of String) = Nothing
                For Each row As DataRow In dttSource.Rows
                    items = row.ItemArray.Select(Function(obj) QuoteValue(obj.ToString(), QuoteValues))
                    writer.WriteLine(String.Join(Delimiter, items))
                Next

                writer.Flush()
            End Using
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function QuoteValue(ByVal value As String, QuoteValues As Boolean) As String
        If QuoteValues = True Then
            Return String.Concat("""", value.Replace("""", """"""), """")
        Else
            Return value
        End If
    End Function

    Public Function DataSetToHtml(dtsInput As DataSet) As String
        If dtsInput Is Nothing Then Return Nothing
        If dtsInput.Tables.Count = 0 Then Return Nothing
        Dim strReturn As String = ""
        For Each tblTable As DataTable In dtsInput.Tables
            strReturn &= DataSetToHtml(tblTable)
            strReturn &= Environment.NewLine
        Next
        Return strReturn
    End Function

    Public Function DataSetToHtml(dttInput As DataTable) As String
        Dim sbrHtml As New StringBuilder
        sbrHtml.AppendLine("<html><head><title>" & dttInput.TableName & "</title></head>")
        sbrHtml.AppendLine("<body><center><table border='1' cellpadding='0' cellspacing='0'>")
        sbrHtml.AppendLine("<tr>")

        For Each dcnCol As DataColumn In dttInput.Columns
            sbrHtml.Append("<td align='center' valign='middle'>" & dcnCol.ColumnName & "</td>")
        Next
        sbrHtml.Append("</tr>")

        For Each drwRow As DataRow In dttInput.Rows
            sbrHtml.AppendLine("<tr>")
            For Each dcnDataCol As DataColumn In dttInput.Columns
                'sbrHtml.Append("<td align='left' valign='middle'>" & If(IsNothing(drwRow(dcnDataCol.ColumnName)), drwRow(dcnDataCol.ColumnName).ToString(), "") & "</td>")
                sbrHtml.Append("<td align='left' valign='middle'>" & drwRow(dcnDataCol.ColumnName) & "</td>")
            Next
            sbrHtml.Append("</tr>")
        Next

        sbrHtml.AppendLine("</table></center></body></html>")

        Return sbrHtml.ToString()
    End Function

    Public Function DeleteFile(FileName As String, Optional Confirm As Boolean = False, Optional Recycle As Boolean = False) As Boolean
        Try
            If CheckFile(FileName) = True Then
                Dim ronRecycle As New FileIO.RecycleOption
                If Recycle = True Then
                    ronRecycle = FileIO.RecycleOption.SendToRecycleBin
                Else
                    ronRecycle = FileIO.RecycleOption.DeletePermanently
                End If
                Dim uioConfirm As New FileIO.UIOption
                If Confirm = True Then
                    uioConfirm = FileIO.UIOption.AllDialogs
                Else
                    uioConfirm = FileIO.UIOption.OnlyErrorDialogs
                End If
                My.Computer.FileSystem.DeleteFile(PathConvert(FileName), uioConfirm, ronRecycle)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region "Search"
    Private _FileRoot As String = "C:\"
    Private _FileFilter As String = "*.*"
    Private _SubFolders As Boolean = False
    Private listForEnumerator As ArrayList

    Public Property FileRoot() As String
        Get
            Return _FileRoot
        End Get
        Set(ByVal Value As String)
            If CheckDir(Value) = True Then
                _FileRoot = Value
            End If
        End Set
    End Property

    Public Property FileFilter() As String
        Get
            Return _FileFilter
        End Get
        Set(ByVal Value As String)
            _FileFilter = Value
        End Set
    End Property

    Public Property SubFolders() As Boolean
        Get
            Return _SubFolders
        End Get
        Set(ByVal Value As Boolean)
            _SubFolders = Value
        End Set
    End Property

    Public Function GetFilesInfo(ByVal strFolderPath As String, ByVal strFileFilter As String) As DataSet
        Dim dataSet1 As New DataSet("DataSet1")
        Dim dteInput As New DataTable("Table1")
        Dim RecordCount As Integer = 0

        Dim Column1 As New DataColumn("FileName")
        Column1.DataType = System.Type.GetType("System.String")
        dteInput.Columns.Add(Column1)

        Dim Column2 As New DataColumn("FileSizeKB")
        Column2.DataType = System.Type.GetType("System.Int32")
        dteInput.Columns.Add(Column2)

        Dim Column3 As New DataColumn("DateCreated")
        Column3.DataType = System.Type.GetType("System.DateTime")
        dteInput.Columns.Add(Column3)

        Dim Column4 As New DataColumn("DateModified")
        Column4.DataType = System.Type.GetType("System.DateTime")
        dteInput.Columns.Add(Column4)

        Dim Column5 As New DataColumn("FileExtension")
        Column5.DataType = System.Type.GetType("System.String")
        dteInput.Columns.Add(Column5)

        Dim Column6 As New DataColumn("FilePath")
        Column6.DataType = System.Type.GetType("System.String")
        dteInput.Columns.Add(Column6)

        Dim Column7 As New DataColumn("ReportDate")
        Column7.DataType = System.Type.GetType("System.DateTime")
        dteInput.Columns.Add(Column7)

        If strFileFilter = "" Then strFileFilter = "*.*"
        dteInput = CrawlFolder(dteInput, strFolderPath, strFileFilter)

        dataSet1.Tables.Add(dteInput)
        GetFilesInfo = dataSet1

    End Function

    Private Function CrawlFolderFast(ByVal dsTable1 As DataTable, ByVal strFolderPath As String, ByVal strFileFilter As String) As DataTable
        Dim csf As New SearchOption
        Dim dirInfo As New IO.DirectoryInfo(strFolderPath)
        If SubFolders = True Then
            csf = SearchOption.AllDirectories
        Else
            csf = SearchOption.TopDirectoryOnly
        End If
        Try
            Dim dirFiles As IO.FileInfo() = dirInfo.GetFiles(strFileFilter, csf)
            Dim dirFile As IO.FileInfo
            For Each dirFile In dirFiles
                Try
                    Dim tRow As DataRow = dsTable1.NewRow
                    dsTable1.Rows.Add(tRow)
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileName") = dirFile.Name
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileSizeKB") = dirFile.Length / 1024
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateCreated") = dirFile.CreationTime
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateModified") = dirFile.LastWriteTime
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileExtension") = dirFile.Extension
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FilePath") = dirFile.DirectoryName
                    dsTable1.Rows(dsTable1.Rows.Count - 1).Item("ReportDate") = Now
                Catch
                    'do nothing
                End Try
            Next
        Catch ex As Exception
            Try
                Dim tRow As DataRow = dsTable1.NewRow
                dsTable1.Rows.Add(tRow)
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileName") = "Error Accessing folder or subfolder " & ex.Message
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileSizeKB") = 0
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateCreated") = Nothing
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateModified") = Nothing
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileExtension") = ""
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FilePath") = strFolderPath
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("ReportDate") = Now
            Catch
                'do nothing
            End Try
        End Try

        Return dsTable1
    End Function

    Private Function CrawlFolder(ByVal dteInput As DataTable, ByVal strFolderPath As String, ByVal strFileFilter As String) As DataTable

        Try
            Dim dirInfo As New IO.DirectoryInfo(strFolderPath)
            dteInput = Crawlfiles(dteInput, dirInfo, strFileFilter)

            If SubFolders = True Then
                Try
                    For Each dioSubDir As DirectoryInfo In dirInfo.GetDirectories
                        Try
                            dteInput = CrawlSubFolder(dteInput, dioSubDir, strFileFilter)
                        Catch ex As Exception
                            Dim tRow As DataRow = dteInput.NewRow
                            dteInput.Rows.Add(tRow)
                            dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing folder or subfolder. " & ex.Message
                            'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = Nothing
                            'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = Nothing
                            'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = Nothing
                            'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = Nothing
                            dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dioSubDir.FullName
                            dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
                            Continue For
                        End Try
                    Next
                Catch ex As Exception
                    Try
                        Dim tRow As DataRow = dteInput.NewRow
                        dteInput.Rows.Add(tRow)
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing main folder subfolders. " & ex.Message
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = 0
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = Nothing
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = Nothing
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = ""
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dirInfo.FullName
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
                    Catch
                        'do nothing
                    End Try
                End Try
            End If
        Catch ex As Exception
            Try
                Dim tRow As DataRow = dteInput.NewRow
                dteInput.Rows.Add(tRow)
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing main folder files. " & ex.Message
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = 0
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = Nothing
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = Nothing
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = ""
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & strFolderPath
                dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
            Catch
                'do nothing
            End Try
        End Try

        Return dteInput
    End Function

    Private Function CrawlSubFolder(ByVal dteInput As DataTable, ByVal dioSubDir As DirectoryInfo, ByVal strFileFilter As String) As DataTable
        Try
            dteInput = Crawlfiles(dteInput, dioSubDir, strFileFilter)
            If SubFolders = True Then
                For Each dioSubSubDir As DirectoryInfo In dioSubDir.GetDirectories
                    Try
                        dteInput = CrawlSubFolder(dteInput, dioSubSubDir, strFileFilter)
                    Catch ex As Exception
                        Dim tRow As DataRow = dteInput.NewRow
                        dteInput.Rows.Add(tRow)
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing subfolder. " & ex.Message
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = 0
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = Nothing
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = Nothing
                        'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = ""
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dioSubSubDir.FullName
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
                        Continue For
                    End Try
                Next
            End If
        Catch ex As Exception
            Try
                Dim tRow As DataRow = dteInput.NewRow
                dteInput.Rows.Add(tRow)
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing sub folder. " & ex.Message
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = 0
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = ""
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = ""
                'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = ""
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dioSubDir.FullName
                dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
            Catch
                'do nothing
            End Try
        End Try

        Return dteInput
    End Function

    Private Function Crawlfiles(ByVal dteInput As DataTable, ByVal dioDir As DirectoryInfo, ByVal strFileFilter As String) As DataTable
        For Each dirFile As IO.FileInfo In dioDir.GetFiles(strFileFilter, SearchOption.TopDirectoryOnly)
            Try
                Dim tRow As DataRow = dteInput.NewRow
                dteInput.Rows.Add(tRow)
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = dirFile.Name
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = dirFile.Length / 1024
                dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = dirFile.CreationTime
                dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = dirFile.LastWriteTime
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = dirFile.Extension
                dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = dirFile.DirectoryName
                dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
            Catch ex As Exception
                Try
                    Dim tRow As DataRow = dteInput.NewRow
                    dteInput.Rows.Add(tRow)
                    dteInput.Rows(dteInput.Rows.Count - 1).Item("FileName") = "Error Accessing file. " & ex.Message
                    'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileSizeKB") = 0
                    'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateCreated") = 0
                    'dteInput.Rows(dteInput.Rows.Count - 1).Item("DateModified") = 0
                    'dteInput.Rows(dteInput.Rows.Count - 1).Item("FileExtension") = ex.GetType().ToString
                    If ex.Message.Contains("The specified path, file name, or both are too long") Then
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dioDir.FullName
                    Else
                        dteInput.Rows(dteInput.Rows.Count - 1).Item("FilePath") = "Error on: " & dirFile.FullName
                    End If
                    dteInput.Rows(dteInput.Rows.Count - 1).Item("ReportDate") = Now
                    Continue For
                Catch
                    Continue For
                End Try
            End Try
        Next
        Return dteInput
    End Function

    Public Function GetFiles(Optional ByVal strFileRoot As String = Nothing, Optional ByVal strFileFilter As String = Nothing) As ArrayList
        ' See if values are passed to the function
        If strFileRoot Is Nothing Then strFileRoot = _FileRoot
        If strFileFilter Is Nothing Then strFileFilter = _FileFilter

        ' Create an array list that will contain the files.
        listForEnumerator = New ArrayList()

        'Populate the arraylist
        GetFilesInFolder(strFileRoot, strFileFilter)
        Return listForEnumerator
    End Function

    Private Sub GetFilesInFolder(ByVal strFileRoot As String, ByVal strFileFilter As String)

        Dim localFiles() As String
        Dim localFile As String
        'Dim fileChangeDate As Date
        'Dim fileAge As TimeSpan
        'Dim fileAgeInDays As Integer
        Dim childFolder As String

        Try
            localFiles = Directory.GetFiles(strFileRoot, strFileFilter)
            For Each localFile In localFiles
                listForEnumerator.Add(localFile)
                'fileChangeDate = File.GetLastWriteTime(localFile)
                'fileAge = DateTime.Now.Subtract(fileChangeDate)
                'fileAgeInDays = fileAge.Days
            Next

            If Directory.GetDirectories(strFileRoot).Length > 0 And SubFolders = True Then
                For Each childFolder In Directory.GetDirectories(strFileRoot)
                    GetFilesInFolder(childFolder, strFileFilter)
                Next
            End If

        Catch
            ' Ignore exceptions on special folders such as System Volume Information.
        End Try

    End Sub

#End Region

#Region "XML"

    Private _XmlDoc As XmlDocument = Nothing

    Public Property XmlDoc() As XmlDocument
        Get
            Return _XmlDoc
        End Get
        Set(ByVal Value As XmlDocument)
            _XmlDoc = Value
        End Set
    End Property

    Public Function CreateRootDocument(ByVal xmlDoc As XmlDocument, ByVal RootNode As String, ByVal FirstNode As String, Optional ByVal StandAlone As Boolean = True) As XmlDocument
        If xmlDoc Is Nothing Then xmlDoc = New XmlDocument
        Dim strStandAlone As String = "no"
        If StandAlone Then strStandAlone = "yes"

        If xmlDoc.FirstChild Is Nothing Then
            Dim xNode As XmlNode = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", strStandAlone)
            xmlDoc.AppendChild(xNode)
        ElseIf xmlDoc.FirstChild.NodeType = XmlNodeType.XmlDeclaration Then
            'Skip this step
        Else
            Dim xNode As XmlNode = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", strStandAlone)
            xmlDoc.InsertBefore(xNode, xmlDoc.FirstChild)
        End If

        If RootNode <> Nothing Then
            If RootNode.Length > 0 Then
                Dim xNode2 As XmlNode = xmlDoc.CreateElement(RootNode)
                xmlDoc.AppendChild(xNode2)
                If FirstNode <> Nothing Then
                    If FirstNode.Length > 0 Then
                        Dim xNode3 As XmlNode = xmlDoc.CreateElement(FirstNode)
                        xNode2.AppendChild(xNode3)
                    End If
                End If
            End If
        End If
        Return xmlDoc
    End Function

    Public Function LoadXml(strPathFile As String) As XmlDocument
        Dim xmlDoc As XmlDocument = CreateRootDocument(Nothing, Nothing, Nothing)
        If CheckFile(PathConvert(strPathFile)) = True Then
            Try
                xmlDoc.Load(PathConvert(strPathFile))
            Catch ex As Exception
                Return Nothing
            End Try
        End If
        Return xmlDoc
    End Function

    Public Function XmlToDataset(xmlDoc As XmlDocument) As DataSet
        Dim rdrXml As New XmlNodeReader(xmlDoc)
        Dim dtsOutput As New DataSet
        dtsOutput.ReadXml(rdrXml)
        Return dtsOutput
    End Function

    Public Function LoadXmlToDataset(strPathFile As String, Optional LoadOnly As Boolean = False) As DataSet
        Dim xmlDoc As XmlDocument = LoadXml(strPathFile)
        _XmlDoc = xmlDoc
        Dim dtsOutput As New DataSet
        If LoadOnly = False Then
            dtsOutput = XmlToDataset(xmlDoc)
        End If
        Return dtsOutput
    End Function

    Public Function LoadXmlFile(ByVal xmlDoc As XmlDocument, ByVal xmlFile As String) As Boolean
        '** This is the first to start. Check to see if the file exists
        If CheckFile(xmlFile) Then
            '** Load the file and check it's integrity
            Try
                xmlDoc.Load(PathConvert(xmlFile))
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function AddNode(ByVal xmlDoc As XmlDocument, ByVal ParentNode As String, ByVal NewNode As XmlNode, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlDocument
        Dim tmpNode As XmlNode = FindXmlNode(xmlDoc, ParentNode, SearchNode, SearchValue)
        tmpNode.AppendChild(NewNode)
        Return xmlDoc
    End Function

    Function CreateAppendElement(ByVal ParentNode As XmlNode, ByVal NodeName As String, Optional ByVal InnerText As String = Nothing, Optional ByVal UpdateMode As Boolean = False) As XmlElement
        If CheckElement(ParentNode, NodeName) = False Or UpdateMode = False Then
            Dim xmlEl As XmlElement = ParentNode.OwnerDocument.CreateElement(NodeName)
            If Not (InnerText Is Nothing) Then xmlEl.InnerText = InnerText
            ParentNode.AppendChild(xmlEl)
            Return xmlEl
        Else
            If Not (InnerText Is Nothing) Then ParentNode.Item(NodeName).InnerText = InnerText
            Return FindXmlNode(ParentNode, NodeName, NodeName, InnerText)
        End If
    End Function

    Function CreateAppendAttribute(ByVal ParentNode As XmlNode, ByVal AttributeName As String, Optional ByVal InnerText As String = Nothing, Optional ByVal UpdateMode As Boolean = True) As XmlElement
        Dim xANode As XmlNode = ParentNode.Attributes.GetNamedItem(AttributeName)
        If UpdateMode = False Or xANode Is Nothing Then
            Dim newAttribute As XmlAttribute = ParentNode.OwnerDocument.CreateAttribute(AttributeName)
            ParentNode.Attributes.Append(newAttribute)
            newAttribute.Value = InnerText
        Else
            ParentNode.Attributes(AttributeName).InnerText = InnerText
        End If
        Return ParentNode
    End Function

    Public Function FindXmlNode(ByVal xmlDoc As XmlDocument, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlNode
        Dim FindNode As XmlNode
        Dim root As XmlElement = xmlDoc.DocumentElement
        If root Is Nothing Then Return Nothing
        'ReturnNode = "//" & ReturnNode
        FindNode = FindXmlNode(root, ReturnNode, SearchNode, SearchValue)
        Return FindNode
    End Function

    Public Function FindXmlNode(ByVal xNode As XmlNode, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlNode
        Dim FindNode As XmlNode
        Dim strXpath As String = ".//" & ReturnNode
        'Dim strReturnNode As String = ReturnNode.Replace("//", "")

        If SearchNode = Nothing Then
            strXpath &= "[1]"
        Else
            If SearchValue = Nothing Then
                If ReturnNode = SearchNode Then
                    strXpath &= "[1]"
                Else
                    strXpath &= "[" & SearchNode & "]"
                End If
            ElseIf ReturnNode = SearchNode Then
                strXpath &= "[text()='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            End If
        End If
        FindNode = xNode.SelectSingleNode(strXpath)
        Return FindNode
    End Function

    Public Function FindXmlNodes(ByVal xmlDoc As XmlDocument, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing, Optional SortField As String = Nothing) As XmlNodeList
        Dim strXpath As String = ReturnNode & ""
        Dim FindNodes As XmlNodeList = Nothing

        Dim root As XmlElement = xmlDoc.DocumentElement
        If root Is Nothing Then Return Nothing

        FindNodes = FindXmlNodes(root, ReturnNode, SearchNode, SearchValue, SortField)
        Return FindNodes
    End Function

    Public Function FindXmlNodes(ByVal xmlDoc As XmlNode, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing, Optional SortField As String = Nothing) As XmlNodeList
        Dim strXpath As String = ".//" & ReturnNode
        Dim FindNodes As XmlNodeList = Nothing

        If Not SearchNode = Nothing Then
            If Not SearchValue = Nothing Then
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "]"
            End If
        End If

        If SortField Is Nothing Then
            FindNodes = xmlDoc.SelectNodes(strXpath)
        Else
            Dim nav As XPathNavigator = xmlDoc.CreateNavigator()
            Dim expr As XPathExpression

            expr = nav.Compile(strXpath)
            expr.AddSort(SortField, XmlSortOrder.Ascending, XmlCaseOrder.None, "", XmlDataType.Number)
            Dim iterator As XPathNodeIterator = nav.Select(expr)

            Dim xmlCDoc As XmlDocument = CreateRootDocument(Nothing, "Sequenchel", Nothing)
            Do While iterator.MoveNext()

                Dim xNode As XmlNode = CType(iterator.Current, IHasXmlNode).GetNode()
                Dim importNode As XmlNode = xmlCDoc.ImportNode(xNode, True)
                xmlCDoc.Item("Sequenchel").AppendChild(importNode)
            Loop

            Dim strReturnNode As String = ReturnNode
            If ReturnNode.LastIndexOf("/") > 0 Then
                strReturnNode = "Sequenchel/" & ReturnNode.Substring(ReturnNode.LastIndexOf("/") + 1, ReturnNode.Length - (ReturnNode.LastIndexOf("/") + 1))
            End If
            FindNodes = xmlCDoc.SelectNodes(strReturnNode)
        End If
        Return FindNodes
    End Function

    Public Function RemoveNode(ByVal xmlDoc As XmlDocument, ByVal OldNode As String, ByVal SearchNode As String, ByVal SearchValue As String) As XmlDocument
        Dim root As XmlElement = xmlDoc.DocumentElement
        If root Is Nothing Then Return xmlDoc
        root = RemoveNode(root, OldNode, SearchNode, SearchValue)
        '** Remove the old node
        'Dim tmpNode As XmlNode = FindXmlNode(xmlDoc, OldNode, SearchNode, SearchValue)
        'If Not tmpNode Is Nothing Then
        '    'xmlDoc.Item("Sequenchel").Item("DataBases").RemoveChild(tmpNode)
        '    tmpNode.ParentNode.RemoveChild(tmpNode)
        'End If
        Return xmlDoc
    End Function

    Public Function RemoveNode(ByVal xmlParentNode As XmlNode, ByVal OldNode As String, ByVal SearchNode As String, ByVal SearchValue As String) As XmlNode
        '** Remove the old node
        Dim tmpNode As XmlNode = FindXmlNode(xmlParentNode, OldNode, SearchNode, SearchValue)
        If Not tmpNode Is Nothing Then
            'xmlDoc.Item("Sequenchel").Item("DataBases").RemoveChild(tmpNode)
            tmpNode.ParentNode.RemoveChild(tmpNode)
        End If
        Return xmlParentNode
    End Function

    Public Sub SaveXmlFile(ByVal xmlDoc As XmlDocument, ByVal FileName As String, Optional ByVal CreateDir As Boolean = False)
        If FileName.Contains("\") Then
            If CheckDir(FileName.Substring(0, FileName.LastIndexOf("\")), CreateDir) = False Then Exit Sub
        End If

        Dim tmpFile As String
        tmpFile = PathConvert(FileName) & ".tmp"
        Dim tmpFileInfo As New FileInfo(tmpFile)
        If tmpFileInfo.Exists = True Then tmpFileInfo.Delete()

        Dim tmpStream As New FileStream(tmpFile, FileMode.Create)
        SaveXmlStream(tmpStream, xmlDoc)

        tmpStream.Close()
        tmpFileInfo.CopyTo(PathConvert(FileName), True)
        tmpFileInfo.Delete()
    End Sub

    Public Sub SaveXmlFile2(ByVal xmlDoc As XmlDocument, ByVal FileName As String, Optional ByVal CreateDir As Boolean = False)
        If FileName.Contains("\") Then
            If CheckDir(FileName.Substring(0, FileName.LastIndexOf("\")), CreateDir) = False Then Exit Sub
        End If

        Try
            Using sw As New System.IO.StringWriter()
                ' Make the XmlTextWriter to format the XML.
                Using xml_writer As New XmlTextWriter(sw)
                    xml_writer.Formatting = Formatting.Indented
                    'dtsInput.WriteXml(xml_writer)
                    xmlDoc.WriteTo(xml_writer)
                    xml_writer.Flush()

                    'Write the XML to disk
                    CreateFile(sw.ToString(), FileName)
                End Using
            End Using

        Catch ex As Exception
            If LogLocation.ToLower = "database" Then
                Dim dhdDb As New DataHandler.db
                dhdDb.WriteLog(ex.Message, 1, LogLevel)
            Else
                WriteLog(ex.Message, 1, LogLevel)
            End If
        End Try
    End Sub

    Public Sub SaveXmlStream(ByVal tmpStream As Stream, ByVal xmlDoc As XmlDocument)
        Dim tmpSerial As New Serialization.XmlSerializer(xmlDoc.GetType)
        tmpSerial.Serialize(tmpStream, xmlDoc)
    End Sub

    Public Function CheckElement(ByVal xmlDoc As XmlDocument, ByVal strName As String) As Boolean
        Dim xNode As XmlNode = FindXmlNode(xmlDoc, strName)
        If xNode Is Nothing Then
            Return False
        End If
        Return True
    End Function

    Public Function CheckElement(ByVal xmlNode As XmlNode, ByVal strName As String) As Boolean
        Dim xNode As XmlNode = FindXmlNode(xmlNode, strName)
        If xNode Is Nothing Then
            Return False
        End If
        Return True
    End Function

    'Friend Function CheckElement(ByVal xmlDoc As XDocument, ByVal name As XName) As Boolean
    '    Return xmlDoc.Descendants(name).Any()
    'End Function

    Public Function LoadItemList(xmlDoc As XmlDocument, strSearchItem As String, strSearchField As String, strSearchValue As String, strTargetItem As String, strDisplayItem As String) As System.Collections.Generic.List(Of String)
        _Errormessage = ""
        Try
            Dim xPNode As System.Xml.XmlNode = FindXmlNode(xmlDoc, strSearchItem, strSearchField, strSearchValue)
            Dim blnSearchValueExists As Boolean = False
            If Not xPNode Is Nothing Then
                Dim ReturnValue As New System.Collections.Generic.List(Of String)
                Dim xNode As System.Xml.XmlNode
                For Each xNode In xPNode.SelectNodes(".//" & strTargetItem)
                    ReturnValue.Add(xNode.Item(strDisplayItem).InnerText)
                Next
                Return ReturnValue
            End If
        Catch ex As Exception
            _Errormessage = ex.Message
        End Try
        Return Nothing
    End Function

    Public Function LoadItemsList(xmlDoc As XmlDocument, strSearchItem As String, strSearchField As String, strSearchValue As String, strTargetItem As String) As System.Collections.Generic.List(Of String)
        _ErrorMessage = ""
        Try
            Dim xNodeList As System.Xml.XmlNodeList = FindXmlNodes(xmlDoc, strSearchItem, strSearchField, strSearchValue)
            Dim blnSearchValueExists As Boolean = False
            If Not xNodeList Is Nothing Then
                Dim ReturnValue As New System.Collections.Generic.List(Of String)
                For Each xNode As System.Xml.XmlNode In xNodeList
                    If CheckElement(xNode, strTargetItem) = True Then
                        ReturnValue.Add(xNode.Item(strTargetItem).InnerText)
                    End If
                Next
                Return ReturnValue
            End If
        Catch ex As Exception
            _ErrorMessage = ex.Message
        End Try
        Return Nothing
    End Function

    Public Sub ExportDataSetToXML(dttInput As DataTable, strFileName As String, Optional ByVal CreateDir As Boolean = False)
        Dim dtsInput As New DataSet
        dtsInput.Tables.Add(dttInput)
        ExportDataSetToXML(dtsInput, strFileName, CreateDir)
    End Sub

    Public Sub ExportDataSetToXML(dtsInput As DataSet, FileName As String, Optional ByVal CreateDir As Boolean = False)
        Try
            Dim xmlDocExport As XmlDocument = CreateRootDocument(Nothing, Nothing, Nothing)
            xmlDocExport.LoadXml(dtsInput.GetXml())
            xmlDocExport = CreateRootDocument(xmlDocExport, Nothing, Nothing)
            SaveXmlFile2(xmlDocExport, FileName, CreateDir)
        Catch ex As Exception
            If LogLocation.ToLower = "database" Then
                Dim dhdDb As New DataHandler.db
                dhdDb.WriteLog(ex.Message, 1, LogLevel)
            Else
                WriteLog(ex.Message, 1, LogLevel)
                'If DevMode Then MessageBox.Show(dhdText.LogFileName & Environment.NewLine & dhdText.LogLocation & Environment.NewLine & dhdText.LogLevel)
            End If
        End Try
    End Sub

#End Region

#Region "Encryption"
    '** Encrypt any string using MD5
    Public Function MD5Encrypt(ByVal strInput As String) As String
        Dim md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim bs As Byte() = md5.ComputeHash(System.Text.Encoding.ASCII.GetBytes(strInput))
        Dim b As Byte
        Dim result As New System.Text.StringBuilder
        For Each b In bs
            result.Append(b.ToString("x2"))
        Next
        Return result.ToString
    End Function

    ' Encrypt the text
    Public Shared Function EncryptText(ByVal Text As String, Optional ByVal ProgKey As String = Nothing) As String
        If ProgKey = Nothing Then ProgKey = "DS%&TS#2"
        Return Encrypt(Text, ProgKey)
    End Function

    'Decrypt the text 
    Public Shared Function DecryptText(ByVal Text As String, Optional ByVal ProgKey As String = Nothing) As String
        If ProgKey = Nothing Then ProgKey = "DS%&TS#2"
        Return Decrypt(Text, ProgKey)
    End Function

    'The function used to encrypt the text
    Private Shared Function Encrypt(ByVal strText As String, ByVal strEncrKey As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}

        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Left(strEncrKey, 8))

            Dim des As New DESCryptoServiceProvider
            Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(strText)
            Dim ms As New MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Convert.ToBase64String(ms.ToArray())

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    'The function used to decrypt the text
    Private Shared Function Decrypt(ByVal strText As String, ByVal sDecrKey As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Dim inputByteArray(strText.Length) As Byte

        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Left(sDecrKey, 8))
            Dim des As New DESCryptoServiceProvider
            inputByteArray = Convert.FromBase64String(strText)
            Dim ms As New MemoryStream
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)

            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8

            Return encoding.GetString(ms.ToArray())

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function
#End Region

#Region "Mail"
    Private _SmtpServer As String = ""
    Private _SmtpCredentials As Boolean = 1
    Private _SmtpUser As String = ""
    Private _SmtpPassword As String = ""
    Private _SmtpReply As String = ""
    Private _SmtpSsl As Boolean = 1
    Private _SmtpPort As Integer = 25
    Private _SmtpRecipient As String = ""

    Public Property SmtpServer() As String
        Get
            Return _SmtpServer
        End Get
        Set(ByVal Value As String)
            _SmtpServer = Value
        End Set
    End Property

    Public Property SmtpCredentials() As Boolean
        Get
            Return _SmtpCredentials
        End Get
        Set(ByVal Value As Boolean)
            _SmtpCredentials = Value
        End Set
    End Property

    Public Property SmtpUser() As String
        Get
            Return _SmtpUser
        End Get
        Set(ByVal Value As String)
            _SmtpUser = Value
        End Set
    End Property

    Public Property SmtpPassword() As String
        Get
            Return _SmtpPassword
        End Get
        Set(ByVal Value As String)
            _SmtpPassword = Value
        End Set
    End Property

    Public Property SmtpReply() As String
        Get
            Return _SmtpReply
        End Get
        Set(ByVal Value As String)
            _SmtpReply = Value
        End Set
    End Property

    Public Property SmtpSsl() As Boolean
        Get
            Return _SmtpSsl
        End Get
        Set(ByVal Value As Boolean)
            _SmtpSsl = Value
        End Set
    End Property

    Public Property SmtpPort() As Integer
        Get
            Return _SmtpPort
        End Get
        Set(ByVal Value As Integer)
            _SmtpPort = Value
        End Set
    End Property

    Public Property SmtpRecipient() As String
        Get
            Return _SmtpRecipient
        End Get
        Set(ByVal Value As String)
            _SmtpRecipient = Value
        End Set
    End Property

    Public Sub SendSMTP(ByVal strFromAddress As String, _
                    ByVal strFromName As String, _
                    ByVal strToAddress As String, _
                    ByVal strToName As String, _
                    ByVal strReplyToAddrr As String, _
                    ByVal strReplyToName As String, _
                    ByVal strSubject As String, _
                    ByVal strBody As String, _
                    ByVal strAttachments As String)

        Dim insMail As New MailMessage(New MailAddress(strFromAddress, strFromName), New MailAddress(strToAddress, strToName))
        If strAttachments = Nothing Then strAttachments = ""
        With insMail
            .Subject = strSubject
            .Body = strBody
            '.ReplyTo = New MailAddress(strReplyToAddrr, strReplyToName)
            .IsBodyHtml = True
            If Not strAttachments.Equals(String.Empty) Then
                Dim strFile As String
                Dim strAttach() As String = strAttachments.Split(";")
                For Each strFile In strAttach
                    .Attachments.Add(New Attachment(strFile.Trim()))
                Next
            End If
        End With

        Dim smtp As New System.Net.Mail.SmtpClient(SmtpServer)
        smtp.EnableSsl = SmtpSsl
        smtp.Port = SmtpPort
        If SmtpCredentials = True Then
            smtp.UseDefaultCredentials = True
        Else
            smtp.UseDefaultCredentials = False
            smtp.Credentials = New System.Net.NetworkCredential(SmtpUser, DecryptText(SmtpPassword))
        End If
        'smtp.Host = CurVar.SmtpServer
        smtp.Send(insMail)
        insMail.Attachments.Dispose()

    End Sub

    Public Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
        Dim emailAddressMatch As RegularExpressions.Match = RegularExpressions.Regex.Match(emailAddress, pattern)
        If emailAddressMatch.Success Then
            EmailAddressCheck = True
        Else
            EmailAddressCheck = False
        End If
        Return EmailAddressCheck
    End Function

#End Region

End Class
