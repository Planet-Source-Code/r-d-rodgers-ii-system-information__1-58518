Attribute VB_Name = "BRPReg"


' Registry API prototypes
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
'Registry Type's
Private Const REG_NONE = 0
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
'Right's for the OpenRegistry
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = ( _
KEY_READ _
And _
KEY_WRITE _
)
Private Const KEY_ALL_ACCESS = ( _
( _
STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or _
KEY_SET_VALUE Or _
KEY_CREATE_SUB_KEY Or _
KEY_ENUMERATE_SUB_KEYS Or _
KEY_NOTIFY Or _
KEY_CREATE_LINK _
) _
And _
( _
Not SYNCHRONIZE _
) _
)
'Local var's to keep track of things hap
'     pening
Dim RootHKey As HKeys
Dim SubDir As String
Dim Hkey As Long
Dim OpenRegOk As Boolean


Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim a As Long
a = RegCreateKey(Hkey, strPath, keyhand)
a = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
a = RegCloseKey(keyhand)
End Sub

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
R = RegOpenKey(Hkey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            getstring = Left$(strBuf, intZeroPos - 1)
        Else
            getstring = strBuf
        End If
    End If
End If
End Function
Public Sub lastboot(hr As TextBox)
 Dim lngHours As Long, lngMinutes As Long
 lngcount = GetTickCount 'boot time in milliseconds
 lngHours = ((lngcount / 1000) / 60) / 60 'full hours since boot
 lngMinutes = ((lngcount / 1000) / 60) Mod 60 'leftover minutes
hr.Text = lngHours & " Hours " & lngMinutes & " Minutes Ago"
End Sub



Function GetAllSubDirectories() As Variant
    On Error GoTo handelgetdirvalues
    Dim SubKey_Num As Integer
    Dim SubKey_Name As String
    Dim Length As Long
    Dim ReturnArray() As Variant
    
    If Not OpenRegOk Then Exit Function
    'Get the Dir List


SubKey_Num = 0


    Do
        Length = 256


SubKey_Name = Space$(Length)


    If RegEnumKey(Hkey, SubKey_Num, SubKey_Name, Length) <> 0 Then
        Exit Do
    End If


SubKey_Name = Left$(SubKey_Name, InStr(SubKey_Name, Chr$(0)) - 1)
    ReDim Preserve ReturnArray(SubKey_Num) As Variant
    ReturnArray(SubKey_Num) = SubKey_Name


SubKey_Num = SubKey_Num + 1
Loop
GetAllSubDirectories = ReturnArray
Exit Function
handelgetdirvalues:
GetAllSubDirectories = Null
Exit Function
End Function


Function GetAllValues() As Variant
    On Error GoTo handelgetdirvalues
    Dim lpData As String, KeyType As Long
    Dim BufferLengh As Long, vname As String, vnamel As Long
    Dim ReturnArray() As Variant, Index As Integer
    
    If Not OpenRegOk Then Exit Function
    
    'Get the Values List
    Index = 0


    Do
        lpData = String(250, " ")
        BufferLengh = 240
        vname = String(250, " ")
        vnamel = 240


        If RegEnumValue(ByVal Hkey, ByVal Index, vname, vnamel, 0, KeyType, lpData, BufferLengh) <> 0 Then
            Exit Do
        End If
        vname = Left$(vname, InStr(vname, Chr$(0)) - 1)
        ReDim Preserve ReturnArray(Index) As Variant
        ReturnArray(Index) = vname
        Index = Index + 1
    Loop
    GetAllValues = ReturnArray
    Exit Function
handelgetdirvalues:
    GetAllValues = Null
    Exit Function
End Function


Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As String
    On Error GoTo handelgetavalue
    Dim i As Integer
    Dim SubKey_Value As String, TempStr As String
    Dim Length As Long
    Dim value_type As Long, RtnVal As Long
    
    If Not OpenRegOk Then Exit Function
    
    'Read the size of the value value
    RtnVal = RegQueryValueEx(Hkey, VarName, 0&, value_type, ByVal 0&, Length)


    Select Case RtnVal
        Case 0 'Ok so continue
        Case 2 'Not Found
        Exit Function
        Case 5 'Access Denied
        GetValue = "Access Denied"
        Exit Function
        Case Else 'What?
        GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
        Exit Function
    End Select
'declare the size of the value and read
'     it


SubKey_Value = Space$(Length)
    RtnVal = RegQueryValueEx(Hkey, VarName, 0&, value_type, ByVal SubKey_Value, Length)


    Select Case value_type
        Case REG_NONE
        'Not defined


SubKey_Value = "Not defined value_type=REG_NONE"
    Case REG_SZ 'A null-terminated string


SubKey_Value = Left$(SubKey_Value, Length - 1)
    Case REG_EXPAND_SZ
SubKey_Value = Left$(SubKey_Value, Length - 1)
    Case REG_BINARY 'Binary data in any form.
SubKey_Value = Left$(SubKey_Value, Length)
    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i


SubKey_Value = TempStr
End If
Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN 'A 32-bit number.


SubKey_Value = Left$(SubKey_Value, Length)


    If Not ReturnBinStr Then
        TempStr = ""


        For i = 1 To Len(SubKey_Value)
            TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
        Next i


SubKey_Value = TempStr
End If
Case REG_DWORD_BIG_ENDIAN
Case REG_LINK
SubKey_Value = "Not defined value_type=REG_LINK"
    Case REG_MULTI_SZ
SubKey_Value = Left$(SubKey_Value, Length)
    Case REG_RESOURCE_LIST
SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"
    Case REG_FULL_RESOURCE_DESCRIPTOR
SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"
    Case REG_RESOURCE_REQUIREMENTS_LIST
SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"
    Case Else
SubKey_Value = "value_type=" & value_type
End Select
GetValue = SubKey_Value
Exit Function
handelgetavalue:
GetValue = ""
Exit Function
End Function
Public Property Get RegistryRootKey() As HKeys
    RegistryRootKey = RootHKey
End Property
Public Property Get SubDirectory() As String
SubDirectory = SubDir
End Property
Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Integer
    On Error GoTo OpenReg
    Dim ReturnVal As Integer


    If RtHKey = 0 Then
        OpenRegistry = False
        OpenRegOk = False
        Exit Function
    End If
    RootHKey = RtHKey
SubDir = SbDr
    If OpenRegOk Then
        CloseRegistry
        OpenRegOk = False
    End If
    ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_READ, Hkey)
    If ReturnVal <> 0 Then
        OpenRegistry = ReturnVal
        Exit Function
    End If
    OpenRegOk = True
    OpenRegistry = True
    Exit Function
OpenReg:
    OpenRegOk = False
    OpenRegistry = False
    Exit Function
End Function
Public Function CloseRegistry() As Boolean
    On Error Resume Next
    If RegCloseKey(Hkey) <> 0 Then
        CloseRegistry = False
        Exit Function
    End If
    CloseRegistry = True
    OpenRegOk = False
End Function
Private Sub Class_Initialize()
RootHKey = &H0
SubDir = ""
    Hkey = 0
    OpenRegOk = False
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    If RegCloseKey(Hkey) <> 0 Then
        Exit Sub
    End If
End Sub

