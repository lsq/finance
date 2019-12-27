Option Explicit


#If Win64 Then
    Public Const PTR_LENGTH As Long = 8
#Else
    Public Const PTR_LENGTH As Long = 4
#End If
 

Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_READ As Long = (( _
                                  STANDARD_RIGHTS_READ _
                                  Or KEY_QUERY_VALUE _
                                  Or KEY_ENUMERATE_SUB_KEYS _
                                  Or KEY_NOTIFY) _
                                  And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Declare PtrSafe Function RegOpenKeyEx _
                          Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
                              ByVal hKey As Long, _
                              ByVal lpSubKey As String, _
                              ByVal ulOptions As Long, _
                              ByVal samDesired As Long, _
                              ByRef phkResult As Long) As Long
Private Declare PtrSafe Function RegEnumKey _
                          Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
                              ByVal hKey As Long, _
                              ByVal dwIndex As Long, _
                              ByVal lpName As String, _
                              ByVal cbName As Long) As Long
Private Declare PtrSafe Function RegQueryValue _
                          Lib "advapi32.dll" Alias "RegQueryValueA" ( _
                              ByVal hKey As Long, _
                              ByVal lpSubKey As String, _
                              ByVal lpValue As String, _
                              ByRef lpcbValue As Long) As Long
Private Declare PtrSafe Function RegCloseKey _
                          Lib "advapi32.dll" ( _
                              ByVal hKey As Long) As Long
Private Declare PtrSafe Function RegOpenKey _
                          Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
                          ByVal hKey As Long, _
                          ByVal lpSubKey As String, _
                          phkResult As Long) As Long
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
                                         (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As _
                                                                                                                     Long, lpData As Any, lpcbData As Long) As Long
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare PtrSafe Function ProgIDFromCLSID Lib "ole32.dll" (ByRef CLSID As GUID, ByRef lplpszProgID As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As LongPtr, ByRef pCLSID As GUID) As Long
Private Declare PtrSafe Function CLSIDFromProgID Lib "ole32" (ByVal lpszProgID As LongPtr, pCLSID As GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef rguid As GUID, ByVal lpstrClsId As LongPtr, ByVal cchMax As Long) As Long
'Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As Long) As Long ' you must use longPtr instead of long
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long ' you must use longPtr instead of long

Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
    
Dim R As Long
Dim arrlist(2 To 10000, 1 To 4) As Variant
Sub TypeLibList()
    ' 重要说明
    ' 首先复制TLBINF32.DLL到C:\Windows\SysWOW64
    ' 必须以管理员角色 regsvr32 C:\Windows\SysWOW64\TLBINF32.DLL
    ' 此时还不能使用的话请参考 https://stackoverflow.com/questions/42569377/tlbinf32-dll-in-a-64bits-net-application
    ' open Windows' "Component Services"
    ' open nodes to "My Computer/COM+ Applications"
    ' right-click, choose to add a new Application
    ' choose an "empty application", name it "tlbinf" for example
    ' make sure you choose "Server application" (means it will be a surrogate that the wizard will be nice to help you create)
    ' choose the user you want the server application to run as (for testing you can choose interactive user but this is an important decision to make)
    ' you don 't have to add any role, not any user
    ' open this newly created app, right-click on "Components" and choose to add a new one
    ' choose to install new component(s)
    ' browse to your tlbinf32.dll location, press "Next" after the wizard has detected 3 interfaces to expose
    ' That 's it. You should see something like this:
    ' enter image description here
    ' Now you can use the same client code and it should work. Note the performance is not comparable however (out-of-process vs in-process).
    ' The surrogate app you've just created has a lots of parameters you can reconfigure later on, with the same UI. You can also script or write code (C#, powershell, VBScript, etc.) to automate all the steps above.
    Dim R1 As Long
    Dim R2 As Long
    Dim hHK1 As Long
    Dim hHK2 As Long
    Dim hHK3 As Long
    Dim hHK4 As Long
    Dim i As Long
    Dim i2 As Long
    
    Dim lpPath As String
    Dim lpGUID As String
    Dim lpName As String
    Dim lpValue As String
    Application.ScreenUpdating = False
    'Cells.Clear: R = 1: Cells(1, 1).Resize(1, 4) = Split("类型库文件路径\类型库引用名称|CLSID|ProgID|默认名称", "|")
    Cells.Clear: R = 1: Cells(1, 1).Resize(1, 4) = Split("类型库文件路径\类型库引用名称|CLSID|ProgID|默认名称", "|")

    lpPath = VBA.String$(128, vbNullChar)
    lpValue = VBA.String$(128, vbNullChar)
    lpName = VBA.String$(128, vbNullChar)
    lpGUID = VBA.String$(128, vbNullChar)
    R1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib", ByVal 0&, KEY_READ, hHK1)
    If R1 = ERROR_SUCCESS Then
        i = 0:
        Do While Not R1 = ERROR_NO_MORE_ITEMS
            R1 = RegEnumKey(hHK1, i, lpGUID, Len(lpGUID))
            If R1 = ERROR_SUCCESS Then
                R2 = RegOpenKeyEx(hHK1, lpGUID, ByVal 0&, KEY_READ, hHK2)
                If R2 = ERROR_SUCCESS Then
                    i2 = 0
                    Do While Not R2 = ERROR_NO_MORE_ITEMS
                        R2 = RegEnumKey(hHK2, i2, lpName, Len(lpName))    '1.0
                        If R2 = ERROR_SUCCESS Then
                            RegQueryValue hHK2, lpName, lpValue, Len(lpValue)
                            RegOpenKeyEx hHK2, lpName, ByVal 0&, KEY_READ, hHK3
                            RegOpenKeyEx hHK3, "0", ByVal 0&, KEY_READ, hHK4
                            RegQueryValue hHK4, "win32", lpPath, Len(lpPath)
                            i2 = i2 + 1
                            'Cells(R + 1, 1) = IIf(InStr(lpPath, vbNullChar), VBA.Left(lpPath, InStr(lpPath, vbNullChar) - 1), lpPath) & VBA.Chr(10) _
                                              & IIf(InStr(lpValue, vbNullChar), VBA.Left(lpValue, InStr(lpValue, vbNullChar) - 1), lpValue) & VBA.Chr(10)
                             arrlist(R + 1, 1) = IIf(InStr(lpPath, vbNullChar), VBA.Left(lpPath, InStr(lpPath, vbNullChar) - 1), lpPath) & VBA.Chr(10) _
                                              & IIf(InStr(lpValue, vbNullChar), VBA.Left(lpValue, InStr(lpValue, vbNullChar) - 1), lpValue) & VBA.Chr(10)

                            ProgIDFromFile lpPath
                        End If
                    Loop
                End If
            End If
            i = i + 1
        Loop
        RegCloseKey hHK1
        RegCloseKey hHK2
        RegCloseKey hHK3
        RegCloseKey hHK4
    End If
    ThisWorkbook.ActiveSheet.Cells(2, 1).Resize(UBound(arrlist), UBound(arrlist, 2)) = arrlist
    Erase arrlist
    Application.ScreenUpdating = True
End Sub

Private Sub ProgIDFromFile(TypeLibFile$)
    Dim CLSID As GUID, strProgID$, lpszProgID As LongPtr
    Dim TLIApp As Object
    Dim TLBInfo As Object
    Dim TypeInf As Object
    'Dim ll As Object
    'Set ll = New TLI.TLIApplication
    Set TLIApp = New TLI.TLIApplication
   ' Set TLIApp = CreateObject("TLI.TLIApplication")
    Dim ProgID As String
    Dim strCLSID As String
    On Error GoTo Exitpoint
    Set TLBInfo = TLIApp.TypeLibInfoFromFile(TypeLibFile)
    For Each TypeInf In TLBInfo.CoClasses
        ProgID = TypeInf.Name
        strCLSID = TypeInf.GUID
        If CLSIDFromString(StrPtr(strCLSID), CLSID) = 0 Then
           ' R = R + 1: Cells(R, 2) = strCLSID
            R = R + 1: arrlist(R, 2) = strCLSID
            'Cells(R, 4) = CLSIDDefaultValue(strCLSID)
            arrlist(R, 4) = CLSIDDefaultValue(strCLSID)
            If ProgIDFromCLSID(CLSID, lpszProgID) = 0 Then
                SysReAllocString VarPtr(strProgID), lpszProgID
                'Mem_Copy strProgID, ByVal lpszProgID, PTR_LENGTH
                'Cells(R, 3) = strProgID
                arrlist(R, 3) = strProgID
                'Cells(R, 3) = CLSIDToProgID(strCLSID)
            End If
        End If
    Next
Exitpoint:

End Sub

Private Function CLSIDDefaultValue(strCLSID$) As String
    Dim ret As Long
    Dim key As Long
    Dim Length As Long
    ret = RegOpenKey(HKEY_CLASSES_ROOT, "CLSID", key)
    ret = RegOpenKey(key, strCLSID, key)
    '先取数据区的长度
    ret = RegQueryValueEx(key, "", 0, 1, ByVal 0, Length)
    '准备数据区
    If Length = 0 Then Exit Function
    Dim buff() As Byte
    ReDim buff(Length - 1)
    '读取数据
    ret = RegQueryValueEx(key, "", 0, 1, buff(0), Length)
    Dim val As String
    '去掉末尾的空字符,VB不需要这个
    ReDim Preserve buff(Length - 2)
    '转化为VB中的字符串
    CLSIDDefaultValue = VBA.StrConv(buff, vbUnicode)
    RegCloseKey (key)
End Function
Public Function ProgIDToCLSID(ByVal strProgID As String) As String
    Dim pCLSDI As GUID
    CLSIDFromProgID StrPtr(strProgID), pCLSID
    ProgIDToCLSID = VBA.String$(38, vbNullChar)
    StringFromGUID2 pCLSID, StrPtr(ProgIDToCLSID), 39
End Function
Public Function CLSIDToProgID(ByVal strCLSID As String) As String
    Dim pCLSID As GUID
    Dim pProgID As LongPtr
    CLSIDFromString StrPtr(strCLSID), pCLSID
    ProgIDFromCLSID pCLSID, pProgID
    SysReAllocString VarPtr(CLSIDToProgID), pProgID
    
End Function

 
' Platform-independent method to return the full zero-padded
' hexadecimal representation of a pointer value
Function HexPtr(ByVal Ptr As LongPtr) As String
    HexPtr = VBA.Hex$(Ptr)
    HexPtr = VBA.String$((PTR_LENGTH * 2) - Len(HexPtr), "0") & HexPtr
End Function
 
Public Function Mem_ReadHex(ByVal Ptr As LongPtr, ByVal Length As Long) As String
    Dim bBuffer() As Byte, strBytes() As String, i As Long, ub As Long, b As Byte
    ub = Length - 1
    ReDim bBuffer(ub)
    ReDim strBytes(ub)
    Mem_Copy bBuffer(0), ByVal Ptr, Length
    For i = 0 To ub
        b = bBuffer(i)
        strBytes(i) = IIf(b < 16, "0", "") & VBA.Hex$(b)
    Next
    Mem_ReadHex = Join(strBytes, "")
End Function

Sub StringPointerExample()
    
    Dim strVar As String, ptrVar As LongPtr, ptrBSTR As LongPtr
    
    strVar = "Hello"
    ptrVar = VarPtr(strVar)
    Mem_Copy ptrBSTR, ByVal ptrVar, PTR_LENGTH
    
    Debug.Print "ptrVar  : 0x"; HexPtr(ptrVar); _
                       " : 0x"; Mem_ReadHex(ptrVar, PTR_LENGTH)
    Debug.Print "ptrBSTR : 0x"; HexPtr(ptrBSTR)
    Debug.Print "StrPtr(): 0x"; HexPtr(StrPtr(strVar))
    Debug.Print "Memory  : 0x"; Mem_ReadHex(ptrBSTR - 4, LenB(strVar) + 6)
    
End Sub
