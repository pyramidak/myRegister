#Region " Register "

Class myRegister

    Public Const Run As String = "Software\Microsoft\Windows\CurrentVersion\Run"

#Region "Procedures "

    Private Shared Function GetRegKey(lngRoot As HKEY) As Microsoft.Win32.RegistryKey
        Select Case CInt(lngRoot)
            Case 0 'HKEY_CLASSES_ROOT
                GetRegKey = Microsoft.Win32.Registry.ClassesRoot
            Case 1 'HKEY_CURRENT_CONFIG
                GetRegKey = Microsoft.Win32.Registry.CurrentConfig
            Case 2 'HKEY_CURRENT_USER
                GetRegKey = Microsoft.Win32.Registry.CurrentUser
            Case 5 'HKEY_PERFORMANCE_DATA
                GetRegKey = Microsoft.Win32.Registry.PerformanceData
            Case 6 'HKEY_USERS
                GetRegKey = Microsoft.Win32.Registry.Users
            Case Else 'HKEY_LOCALE_MACHINE = 4
                GetRegKey = Microsoft.Win32.Registry.LocalMachine
        End Select
    End Function

    Public Shared Function DoesKeyExist(lngRootKey As HKEY, strKey As String) As Boolean
        Dim objRegKey As Microsoft.Win32.RegistryKey
        Dim bOK As Boolean

        objRegKey = GetRegKey(lngRootKey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey, False)
        Catch
        End Try

        If objRegKey Is Nothing Then
            bOK = False
        Else
            bOK = True
        End If
        If Not objRegKey Is Nothing Then
            objRegKey.Close()
            objRegKey = Nothing
        End If

        DoesKeyExist = bOK
    End Function
#End Region

#Region "Create "

    Public Shared Function CreateKey(lngrootkey As HKEY, strKey As String) As Boolean
        CreateKey = False
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.CreateSubKey(strKey)
            If objRegKey IsNot Nothing Then
                CreateKey = True
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function

    Public Shared Function CreateValue(lngrootkey As HKEY, strKey As String, strValName As String, objVal As Object) As Boolean
        CreateValue = False
        If DoesKeyExist(lngrootkey, strKey) = False Then
            If CreateKey(lngrootkey, strKey) = False Then Exit Function
        End If

        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey, True)
            If objRegKey IsNot Nothing Then
                objRegKey.SetValue(strValName, objVal, If(IsNumeric(objVal), Microsoft.Win32.RegistryValueKind.DWord, Microsoft.Win32.RegistryValueKind.String))
                objRegKey.Flush()
                CreateValue = True
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function

#End Region

#Region "Delete "

    Public Shared Function DeleteKey(lngrootkey As HKEY, strKey As String, Optional bRecursive As Boolean = False) As Boolean
        DeleteKey = False
        If DoesKeyExist(lngrootkey, strKey) = False Then Exit Function
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            If objRegKey.OpenSubKey(strKey).SubKeyCount > 0 Then
                objRegKey.DeleteSubKeyTree(strKey)
                DeleteKey = True
            Else
                objRegKey.DeleteSubKey(strKey)
                DeleteKey = True
            End If
            objRegKey.Close()
        Catch
        End Try
    End Function

    ' Registry-Wert löschen (Schlüssel muss existieren) 
    Public Shared Function DeleteValue(lngrootkey As HKEY, strKey As String, strValName As String) As Boolean
        DeleteValue = False
        If strValName = "" Then Exit Function
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey, True)
            If objRegKey IsNot Nothing Then
                objRegKey.DeleteValue(strValName)
                objRegKey.Flush()
                DeleteValue = True
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function
#End Region

#Region "Query "

    Public Shared Function GetValue(lngrootkey As HKEY, strKey As String, strValName As String, objDefault As String) As String
        Return QueryValue(lngrootkey, strKey, strValName, objDefault).ToString
    End Function

    Public Shared Function GetValue(lngrootkey As HKEY, strKey As String, strValName As String, objDefault As Integer) As Integer
        Dim obj As Object = QueryValue(lngrootkey, strKey, strValName, objDefault)
        If IsNumeric(obj) Then
            Return CInt(obj)
        Else
            Return objDefault
        End If
    End Function

    Public Shared Function GetValue(lngrootkey As HKEY, strKey As String, strValName As String, objDefault As Boolean) As Boolean
        Dim obj As Object = QueryValue(lngrootkey, strKey, strValName, objDefault)
        If IsNumeric(obj) AndAlso CInt(obj) = 0 Or CInt(obj) = 1 Then
            Return CBool(obj)
        Else
            Return objDefault
        End If
    End Function

    Private Shared Function QueryValue(lngrootkey As HKEY, strKey As String, strValName As String, objDefault As Object) As Object
        QueryValue = objDefault
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey)
            If objRegKey IsNot Nothing Then
                QueryValue = objRegKey.GetValue(strValName, objDefault)
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function

    Public Shared Function QueryNames(lngrootkey As HKEY, strKey As String) As String()
        Dim Nic(0) As String
        QueryNames = Nic
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey)
            If objRegKey IsNot Nothing Then
                QueryNames = objRegKey.GetValueNames()
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function

    Public Shared Function QueryKeys(lngrootkey As HKEY, strKey As String) As String()
        Dim Nic(0) As String
        QueryKeys = Nic
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        Try
            objRegKey = objRegKey.OpenSubKey(strKey)
            If objRegKey IsNot Nothing Then
                QueryKeys = objRegKey.GetSubKeyNames()
                objRegKey.Close()
            End If
        Catch
        End Try
    End Function
#End Region

#Region "Find "

    Public Shared Function FindValue(lngrootkey As HKEY, strKey As String, strVal As String) As String
        Dim objRegKey As Microsoft.Win32.RegistryKey

        objRegKey = GetRegKey(lngrootkey)
        FindValue = ""
        Try
            objRegKey = objRegKey.OpenSubKey(strKey)
            If objRegKey Is Nothing Then
                Exit Function
            Else
                For Each oneValueName As String In objRegKey.GetValueNames
                    Dim defValue As String = ""
                    If LCase(CStr(objRegKey.GetValue(oneValueName, defValue))).Contains(LCase(strVal)) Then
                        FindValue = oneValueName
                        Exit For
                    End If
                Next
            End If
            objRegKey.Close()
        Catch
        End Try
    End Function
#End Region

#Region "MyApp "

    Public Shared Function GetCloudMyApp() As Integer
        Return CType(myRegister.GetValue(HKEY.CURRENT_USER, "Software\pyramidak\" & Application.ExeName, "Cloud", 0), Integer)
    End Function

    Public Shared Sub WriteCloudMyApp(iCloud As Integer)
        myRegister.CreateValue(HKEY.CURRENT_USER, "Software\pyramidak\" & Application.ExeName, "Cloud", iCloud)
    End Sub

    Public Shared Function GetAutoUpdateMyApp() As Boolean
        Return myRegister.GetValue(HKEY.CURRENT_USER, "Software\pyramidak\" & Application.ExeName, "AutoUpdate", True)
    End Function

    Public Shared Sub WriteAutoUpdateMyApp(doupdate As Boolean)
        myRegister.CreateValue(HKEY.CURRENT_USER, "Software\pyramidak\" & Application.ExeName, "AutoUpdate", doupdate)
    End Sub

    Public Shared Function GetAutoStartMyApp() As Boolean
        Return If(myRegister.GetValue(HKEY.CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Application.ExeName, "").ToLower = Chr(34) & System.Reflection.Assembly.GetExecutingAssembly().Location.ToLower & Chr(34) & " -win", True, False)
    End Function

    Public Shared Sub WriteAutoStartMyApp()
        myRegister.CreateValue(HKEY.CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Application.ExeName, Chr(34) & System.Reflection.Assembly.GetExecutingAssembly().Location & Chr(34) & " -win")
    End Sub

    Public Shared Sub DeleteAutoStartMyApp()
        myRegister.DeleteValue(HKEY.CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", Application.ExeName)
    End Sub

#End Region

End Class

#End Region