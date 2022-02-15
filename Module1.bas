Attribute VB_Name = "Module1"
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Declare Function ExitWindowsEx Lib "user32" (ByVal _
uFlags As Long, ByVal dwReserved As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias _
"RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
Hkey As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal Reserved As Long, ByVal dwType As Long, _
lpData As Any, ByVal cbData As Long) As Long

Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Public Sub savestring(Hkey As Long, strPath As String, _
strValue As String, strdata As String)
Dim keyhand
Dim r
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal _
strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub


