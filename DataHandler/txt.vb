Imports System.IO
Imports System
Imports System.Data
Imports System.Math
Imports System.Collections
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Text
Imports System.Environment.SpecialFolder

Public Class txt

#Region "General"
    Private _Errormessage As String = ""

    Public ReadOnly Property Errormessage() As String
        Get
            Return _Errormessage
        End Get
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

    Public Sub WriteLog(ByVal LogText As String, ByVal EntryLevel As Integer, Optional ByVal Sender As String = Nothing)
        Dim booLogItem As Boolean = False
        If Sender = Nothing Then Sender = Environment.MachineName

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
    End Sub

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
        If FileName.Length < 2 Then Return False

        Dim intCheckDir As Integer = 0, strCheckdir As String = Nothing
        intCheckDir = FileName.IndexOf("\")
        If intCheckDir > 0 Then
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
        Dim dsTable1 As New DataTable("Table1")
        Dim RecordCount As Integer = 0

        Dim Column1 As New DataColumn("FileName")
        Column1.DataType = System.Type.GetType("System.String")
        dsTable1.Columns.Add(Column1)

        Dim Column2 As New DataColumn("FileSizeKB")
        Column2.DataType = System.Type.GetType("System.Int32")
        dsTable1.Columns.Add(Column2)

        Dim Column3 As New DataColumn("DateCreated")
        Column3.DataType = System.Type.GetType("System.DateTime")
        dsTable1.Columns.Add(Column3)

        Dim Column4 As New DataColumn("DateModified")
        Column4.DataType = System.Type.GetType("System.DateTime")
        dsTable1.Columns.Add(Column4)

        Dim Column5 As New DataColumn("FileExtension")
        Column5.DataType = System.Type.GetType("System.String")
        dsTable1.Columns.Add(Column5)

        Dim Column6 As New DataColumn("FilePath")
        Column6.DataType = System.Type.GetType("System.String")
        dsTable1.Columns.Add(Column6)

        Dim Column7 As New DataColumn("ReportDate")
        Column7.DataType = System.Type.GetType("System.DateTime")
        dsTable1.Columns.Add(Column7)

        If strFileFilter = "" Then strFileFilter = "*.*"
        CrawlFolder(dsTable1, strFolderPath, strFileFilter)

        dataSet1.Tables.Add(dsTable1)
        GetFilesInfo = dataSet1

    End Function

    Private Function CrawlFolder(ByVal dsTable1 As DataTable, ByVal strFolderPath As String, ByVal strFileFilter As String) As DataTable
        Dim csf As New SearchOption
        Dim Reportdate As Date = Now()
        Dim dirInfo As New IO.DirectoryInfo(strFolderPath)
        If SubFolders = True Then
            csf = SearchOption.AllDirectories
        Else
            csf = SearchOption.TopDirectoryOnly
        End If
        Dim dirFiles As IO.FileInfo() = dirInfo.GetFiles(strFileFilter, csf)
        Dim dirFile As IO.FileInfo

        For Each dirFile In dirFiles
            Dim tRow As DataRow = dsTable1.NewRow
            dsTable1.Rows.Add(tRow)

            Try
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileName") = dirFile.Name
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileSizeKB") = dirFile.Length / 1024
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateCreated") = dirFile.CreationTime
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("DateModified") = dirFile.LastWriteTime
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FileExtension") = dirFile.Extension
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("FilePath") = dirFile.DirectoryName
                dsTable1.Rows(dsTable1.Rows.Count - 1).Item("ReportDate") = Reportdate
            Catch ex As Exception
                Return Nothing
            End Try
        Next

        'If Directory.GetDirectories(dirInfo.FullName).Length > 0 Then
        '    For Each childFolder As String In Directory.GetDirectories(dirInfo.FullName)
        '        CrawlFolder(dsTable1, childFolder, strFileFilter)
        '    Next
        'End If

        Return dsTable1
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

            If Directory.GetDirectories(strFileRoot).Length > 0 Then
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
            xmlDoc.AppendChild(xNode)
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

    Public Function AddNode(ByVal xmldoc As XmlDocument, ByVal ParentNode As String, ByVal NewNode As XmlNode, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlDocument
        Dim tmpNode As XmlNode = FindXmlNode(xmldoc, ParentNode, SearchNode, SearchValue)
        tmpNode.AppendChild(NewNode)
        Return xmldoc
    End Function

    Function CreateAppendElement(ByVal ParentNode As XmlNode, ByVal NodeName As String, Optional ByVal InnerText As String = Nothing, Optional ByVal UpdateMode As Boolean = False) As XmlElement
        If CheckNodeElement(ParentNode, NodeName) = False Or UpdateMode = False Then
            Dim xmlEl As XmlElement = ParentNode.OwnerDocument.CreateElement(NodeName)
            If Not (InnerText Is Nothing) Then xmlEl.InnerText = InnerText
            ParentNode.AppendChild(xmlEl)
            Return xmlEl
        Else
            If Not (InnerText Is Nothing) Then ParentNode.Item(NodeName).InnerText = InnerText
            Return FindXmlChildNode(ParentNode, NodeName, NodeName, InnerText)
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

        Dim strXpath As String = "//" & ReturnNode
        If SearchNode = Nothing Then
            strXpath &= "[1]"
        Else
            If SearchValue = Nothing Then
                strXpath &= "[" & SearchNode & "]"
            ElseIf ReturnNode = SearchNode Then
                strXpath &= "[text()='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            End If
        End If
        FindNode = root.SelectSingleNode(strXpath)
        Return FindNode
    End Function

    Public Function FindXmlChildNode(ByVal xmlDoc As XmlNode, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlNode
        Dim FindNode As XmlNode
        Dim strXpath As String = ReturnNode
        If SearchNode = Nothing Then
            strXpath &= "[1]"
        Else
            If SearchValue = Nothing Then
                strXpath &= "[" & SearchNode & "]"
            ElseIf ReturnNode = SearchNode Then
                strXpath &= "[text()='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            End If
        End If
        FindNode = xmlDoc.SelectSingleNode(strXpath)
        Return FindNode
    End Function

    Public Function FindXmlNodes(ByVal xmlDoc As XmlDocument, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlNodeList
        Dim FindNodes As XmlNodeList
        Dim root As XmlElement = xmlDoc.DocumentElement
        If root Is Nothing Then Return Nothing
        Dim strXpath As String = ReturnNode & ""
        If Not SearchNode = Nothing Then
            If Not SearchValue = Nothing Then
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "]"
            End If
        End If
        FindNodes = root.SelectNodes(strXpath)
        Return FindNodes
    End Function

    Public Function FindXmlChildNodes(ByVal xmlDoc As XmlNode, ByVal ReturnNode As String, Optional ByVal SearchNode As String = Nothing, Optional ByVal SearchValue As String = Nothing) As XmlNodeList
        Dim FindNodes As XmlNodeList
        Dim strXpath As String = ReturnNode & ""
        If Not SearchNode = Nothing Then
            If Not SearchValue = Nothing Then
                strXpath &= "[" & SearchNode & "='" & SearchValue & "']"
            Else
                strXpath &= "[" & SearchNode & "]"
            End If
        End If
        FindNodes = xmlDoc.SelectNodes(strXpath)
        Return FindNodes
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

    Public Function RemoveNode(ByVal xmlDoc As XmlDocument, ByVal OldNode As String, ByVal SearchNode As String, ByVal SearchValue As String) As XmlDocument
        '** Remove the old node
        Dim tmpNode As XmlNode = FindXmlNode(xmlDoc, OldNode, SearchNode, SearchValue)
        If Not tmpNode Is Nothing Then
            'xmlDoc.Item("Sequenchel").Item("DataBases").RemoveChild(tmpNode)
            tmpNode.ParentNode.RemoveChild(tmpNode)
        End If
        Return xmlDoc
    End Function

    Public Sub SaveXmlFile(ByVal xmlDoc As XmlDocument, ByVal FileName As String)
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

    Public Function CheckNodeElement(ByVal xmlInput As XmlNode, ByVal strName As String) As Boolean
        Dim xNode As XmlNode = FindXmlChildNode(xmlInput, strName)
        If xNode Is Nothing Then
            Return False
        End If
        Return True
    End Function

    Public Function LoadItemList(xmlDoc As XmlDocument, strSearchItem As String, strSearchField As String, strSearchValue As String, strTargetItem As String, strDisplayItem As String) As System.Collections.Generic.List(Of String)
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
        Return Nothing
    End Function

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

End Class
