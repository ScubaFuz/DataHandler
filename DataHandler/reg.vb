Public Class reg

    Private _ErrorLevel As Integer = 0
    Private _MessageRegistryError As String = "Error accessing the registry. Please check your permissions"
    Private _RegMessage As String = ""
    Private _RegistryPath As String = ""
    Private _RegLocation As String = ""

    Public ReadOnly Property ErrorLevel() As Integer
        Get
            Return _ErrorLevel
        End Get
    End Property

    Public ReadOnly Property RegMessage() As String
        Get
            Return _RegMessage
        End Get
    End Property

    Public Property RegistryPath() As String
        Get
            Return _RegistryPath
        End Get
        Set(ByVal Value As String)
            _RegistryPath = Value
        End Set
    End Property

    Public ReadOnly Property RegLocation() As String
        Get
            Return _RegLocation
        End Get
    End Property

    Public Function AddLMRegKey(ByVal keyName As String, ByVal keyValue As Object, Optional ByVal RegPath As String = Nothing) As Integer
        _RegMessage = ""
        Dim key As Microsoft.Win32.RegistryKey
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(RegPath)
            key.SetValue(keyName, keyValue)
            _ErrorLevel = 0
            _RegLocation = "HKLM"
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
        Finally
            Microsoft.Win32.Registry.LocalMachine.Close()
        End Try
        Return _ErrorLevel
    End Function

    Public Function AddCURRegKey(ByVal keyName As String, ByVal keyValue As Object, Optional ByVal RegPath As String = Nothing) As Integer
        _RegMessage = ""
        Dim key As Microsoft.Win32.RegistryKey
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegPath)
            key.SetValue(keyName, keyValue)
            _ErrorLevel = 0
            _RegLocation = "HKCU"
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
        Finally
            Microsoft.Win32.Registry.CurrentUser.Close()
        End Try
        Return _ErrorLevel
    End Function

    Public Function AddAnyRegKey(ByVal keyName As String, ByVal keyValue As Object, Optional ByVal RegPath As String = Nothing) As Integer
        Try
            AddLMRegKey(keyName, keyValue, RegPath)
            If ErrorLevel = -1 Then
                AddCURRegKey(keyName, keyValue, RegPath)
            End If
            'If ErrorLevel = -1 Then
            '    'ohoh
            'End If
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
        End Try
        Return _ErrorLevel
    End Function

    Public Function ReadLMRegKey(ByVal keyName As String, Optional ByVal RegPath As String = Nothing) As String
        _RegMessage = ""
        Dim key As Microsoft.Win32.RegistryKey
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            _RegLocation = "HKLM"
            key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(RegPath)
            Dim keyValue As String = CType(key.GetValue(keyName), String)
            _ErrorLevel = 0
            Return keyValue
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
            Return _ErrorLevel
        Finally
            Microsoft.Win32.Registry.LocalMachine.Close()
        End Try
    End Function

    Public Function ReadCURRegKey(ByVal keyName As String, Optional ByVal RegPath As String = Nothing) As String
        _RegMessage = ""
        Dim key As Microsoft.Win32.RegistryKey
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            _RegLocation = "HKCU"
            key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegPath)
            Dim keyValue As String = CType(key.GetValue(keyName), String)
            _ErrorLevel = 0
            Return keyValue
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
            Return _ErrorLevel
        Finally
            Microsoft.Win32.Registry.CurrentUser.Close()
        End Try
    End Function

    Public Function ReadAnyRegKey(ByVal keyName As String, Optional ByVal RegPath As String = Nothing) As String
        Dim strValue As String = Nothing
        Try
            strValue = ReadCURRegKey(keyName, RegPath)
            If strValue = "-1" Then strValue = ReadLMRegKey(keyName, RegPath)
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
            strValue = "-1"
        End Try
        Return strValue
    End Function

    Public Function DeleteLMRegKey(ByVal keyName As String, Optional ByVal RegPath As String = Nothing) As String
        _RegMessage = ""
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            _RegLocation = "HKLM"
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(RegPath & "\" & keyName)
            _ErrorLevel = 0
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
        Finally
            Microsoft.Win32.Registry.CurrentUser.Close()
        End Try
        Return _ErrorLevel
    End Function

    Public Function DeleteCURRegKey(ByVal keyName As String, Optional ByVal RegPath As String = Nothing) As String
        _RegMessage = ""
        If RegPath = Nothing Then RegPath = _RegistryPath
        Try
            _RegLocation = "HKCU"
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(RegPath & "\" & keyName)
            _ErrorLevel = 0
        Catch ex As Exception
            _RegMessage = _MessageRegistryError & "  Path: " & _RegLocation & "\" & RegPath & vbCrLf & ex.Message
            _ErrorLevel = -1
        Finally
            Microsoft.Win32.Registry.CurrentUser.Close()
        End Try
        Return _ErrorLevel
    End Function

End Class
