Attribute VB_Name = "Module1"
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY _
Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) _
And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
    
Declare Sub PwdChangePassword Lib "mpr.dll" Alias _
    "PwdChangePasswordA" (ByVal lpcRegkeyname As String, _
    ByVal hwnd As Long, ByVal uireserved1 As Long, ByVal _
    uireserved2 As Long)
    
Declare Function VerifyScreenSavePwd Lib _
    "password.cpl" (ByVal hwnd As Long) As Boolean
    
Public Sub SetLoaded()
    'put this in your main forms' Load proce
    '     dure
    'this will set the count
    Dim lTemp As Long, sPath As String
    lTemp& = GetLoaded&
    If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & "wohsspass.ini" Else sPath$ = App.Path & "wohsspass.ini"
    Open sPath$ For Output As #1
    Print #1, lTemp& + 1
    Close #1
End Sub


Public Function GetLoaded() As Long
    'call this to get how many times program
    '     has been loaded
    On Error Resume Next
    Dim sPath As String, sTemp As String
    If Right$(App.Path, 1) <> "\" Then sPath$ = App.Path & "\" & "wohsspass.ini" Else sPath$ = App.Path & "wohsspass.ini"
    Open sPath$ For Input As #1
    sTemp$ = Input(LOF(1), #1)
    Close #1
    If sTemp$ = "" Then GetLoaded& = 0 Else GetLoaded& = CLng(sTemp$)
End Function


Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


Function sdaGetRegEntry(strKey As String, _
    strSubKeys As String, strValName As String, _
    lngType As Long) As String
    '* Demonstration of win32 API's to query
    '
    ' the system registry
    ' Stu Alderman -- 2/30/96
    On Error GoTo sdaGetRegEntry_Err
    Dim lngResult As Long, lngKey As Long
    Dim lngHandle As Long, lngcbData As Long
    Dim strRet As String


    Select Case strKey
        Case "HKEY_CLASSES_ROOT": lngKey = &H80000000
        Case "HKEY_CURRENT_CONFIG": lngKey = &H80000005
        Case "HKEY_CURRENT_USER": lngKey = &H80000001
        Case "HKEY_DYN_DATA": lngKey = &H80000006
        Case "HKEY_LOCAL_MACHINE": lngKey = &H80000002
        Case "HKEY_PERFORMANCE_DATA": lngKey = &H80000004
        Case "HKEY_USERS": lngKey = &H80000003
        Case Else: Exit Function
    End Select
If Not ERROR_SUCCESS = RegOpenKeyEx(lngKey, _
strSubKeys, 0&, KEY_READ, _
lngHandle) Then Exit Function
lngResult = RegQueryValueEx(lngHandle, strValName, _
0&, lngType, ByVal strRet, lngcbData)
strRet = Space(lngcbData)
lngResult = RegQueryValueEx(lngHandle, strValName, _
0&, lngType, ByVal strRet, lngcbData)
If Not ERROR_SUCCESS = RegCloseKey(lngHandle) Then _
lngType = -1&
sdaGetRegEntry = strRet
sdaGetRegEntry_Exit:
On Error GoTo 0
Exit Function
sdaGetRegEntry_Err:
lngType = -1&
MsgBox Err & "> " & Error$, 16, _
"GenUtils/sdaGetRegEntry"
Resume sdaGetRegEntry_Exit
End Function
'End Function


