Attribute VB_Name = "Module1"
Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Function GetFromFile(ByVal strFile, ByVal strSection, ByVal strKey, ByVal strDefault As String)
    Dim Length As Long
    Dim retval As String
    retval = Space(255)
    Length = GetPrivateProfileString(strSection, strKey, strDefault, retval, 255, strFile)
    retval = Left(retval, Length)
    GetFromFile = retval
End Function

Function WriteToFile(ByVal File, ByVal Section, ByVal Key, ByVal Data)
    Dim retval As Long
    retval = WritePrivateProfileString(Section, Key, Data, File)
End Function

