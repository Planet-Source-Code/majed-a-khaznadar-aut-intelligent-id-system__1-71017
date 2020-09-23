Attribute VB_Name = "INI"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal Filename As String) As Long

Public Function READINI(inifile, inisection, inikey, iniDefault)
'Fail fracefully if no file / wrong file is specified.
'If no section (appname), default is first appname
'if no key, default is first key


Dim lpApplicationName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpReturnedString As String
Dim nSize As Long
Dim lpFileName As String
Dim retval As Long
Dim Filename As String
lpDefault = Space$(254)
lpDefault = iniDefault

lpReturnedString = Space$(254)

nSize = 254
lpFileName = inifile
lpApplicationName = inisection
lpKeyName = inikey
Filename = lpFileName
retval = GetPrivateProfileString _
(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
READINI = lpReturnedString
End Function


Public Function WRITEINI(inifile As String, inisection As String, inikey As String, Info As String) As String
Dim retval As Long
retval = WritePrivateProfileString(inisection, inikey, Info, inifile)
WRITEINI = LTrim$(Str$(retval))
End Function



