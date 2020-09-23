Attribute VB_Name = "mdlRewrite"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public varTmp As Variant, varTwo As Variant, strFree As String, strFile As String, e%, i%

Public Function ReadINI(Section, KeyName, filename As String) As String
  Dim sRet As String
  sRet = String(998, Chr(0))
  ReadINI = Left(sRet, GetPrivateProfileString(Section, KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(Section, KeyName, NewString As String, filename As String) As String
  Dim sWet As String
  sWet = WritePrivateProfileString(Section, KeyName, NewString, filename)
End Function

Public Sub ListToList(lstOne As ListBox, lstTmp As ListBox)
  For i% = 0 To lstTmp.ListCount - 1
    lstOne.AddItem lstTmp.List(i%)
  Next i%
End Sub

Public Sub FindAndRemove(lstBox As ListBox, strFind As String)
  For i% = 0 To lstBox.ListCount - 1
    If lstBox.List(i%) = strFind Then lstBox.RemoveItem i%: Exit Sub
  Next i%
End Sub

Public Sub SplitAndAdd(objBox As Object, varAdd As Variant)
  Dim varOpt As Variant
  varOpt = Split(varAdd, "")
  For i% = 1 To UBound(varOpt)
    objBox.AddItem varOpt(i%)
  Next i%
End Sub
