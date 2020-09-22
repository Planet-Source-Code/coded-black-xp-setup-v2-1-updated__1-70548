VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mIniFileName As String
Public valueygdibaca As String

Public Property Let FileName(ByVal INIFileName As String)
Attribute FileName.VB_Description = "Set or Get the ini file name,required for working with the file"
    If Dir(INIFileName, vbNormal) = "" Then
      Open INIFileName For Output As #1
      Close #1
    End If
    mIniFileName = INIFileName
End Property

Public Property Get FileName() As String
    FileName = mIniFileName
End Property

Public Function GetValue(ByVal Section As String, ByVal Key As String, Optional ByVal DefaultValue As String) As String
Attribute GetValue.VB_Description = "retrive the value of a key in a section"
  On Error GoTo waduh_salah
  Dim Value As String, X As Long
  valueygdibaca = String$(500, 0)
  X = GetPrivateProfileString(Section, Key, DefaultValue, valueygdibaca, Len(valueygdibaca), mIniFileName)
  GetValue = Trim(Left(valueygdibaca, X))
  hasilbacavalue = GetValue
Exit Function
waduh_salah:
  GetValue = DefaultValue
End Function

Public Function WriteValue(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
Attribute WriteValue.VB_Description = "Write the valueto the ini file"
  On Error GoTo waduh_salah
  Dim X As Integer
  X = WritePrivateProfileString(Section, Key, Value, mIniFileName)
  If X <> 0 Then WriteValue = True
  Exit Function
waduh_salah:
End Function

Public Function GetAllSections() As Collection
Attribute GetAllSections.VB_Description = "Retrive the list of section in the ini file"
  Dim Value As String, retval As String, X As Integer
  Dim S() As String, i As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(vbNullString, "", "", retval, Len(retval), mIniFileName)
  Value = Trim(Left(retval, X))
  S = Split(Value, Chr(0))
  Set GetAllSections = New Collection
  With GetAllSections
    For i = LBound(S) To UBound(S)
      If S(i) <> "" Then .Add S(i)
    Next
  End With
End Function

Public Function GetAllKeys(ByVal Section As String) As Collection
Attribute GetAllKeys.VB_Description = "retrive all the key in the specified section"
  Dim Value As String, retval As String, X As Integer
  Dim S() As String, i As Integer
  retval = String$(255, 0)
  X = GetPrivateProfileString(Section, vbNullString, "", retval, Len(retval), mIniFileName)
  Value = Trim(Left(retval, X))
  S = Split(Value, Chr(0))
  Set GetAllKeys = New Collection
  With GetAllKeys
    For i = LBound(S) To UBound(S)
      If S(i) <> "" Then .Add S(i)
    Next
  End With
End Function

Public Function DeleteSection(ByVal Section As String) As Boolean
Attribute DeleteSection.VB_Description = "Delete all the keys in the specified section"
  On Error GoTo waduh_salah
  Dim X As Integer
  X = WritePrivateProfileString(Section, vbNullString, "", mIniFileName)
  If X <> 0 Then DeleteSection = True
  Exit Function
waduh_salah:
End Function

Public Function DeleteKey(ByVal Section As String, ByVal Key As String) As Boolean
Attribute DeleteKey.VB_Description = "Delete key in the ini file "
  On Error GoTo waduh_salah
  Dim X As Integer
  X = WritePrivateProfileString(Section, Key, vbNullString, mIniFileName)
  If X <> 0 Then DeleteKey = True
  Exit Function
waduh_salah:
End Function
