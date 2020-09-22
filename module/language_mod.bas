Attribute VB_Name = "language_mod"
Option Explicit

Public lang As String
Public plhn As Integer
Public ini_baru As Boolean
Public ayo_load_osh As String
Public ayo_load_lang As String
Public hasilbacavalue As String

Public Sub detect_lang()
ini_baru = Dir(App.Path & "\data\bin\xp-setup.ini") = ""
ini_file.FileName = App.Path & "\data\bin\xp-setup.ini"
With ini_file
    If ini_baru Then
        .WriteValue "Language", "lang", "En"
    Else
        .GetValue "Language", "lang"
    End If
    If Not hasilbacavalue = "En" And Not hasilbacavalue = "Id" Then
        .WriteValue "Language", "lang", "En"
    End If
End With
home_set_lang
End Sub

Public Function load_lang(nama_frm As String, index_frm As Form, jmlh_label As Integer, Optional fungsi_ato_tidak As Boolean) As String
Dim a As Integer
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
With ini_file
    For a = 0 To jmlh_label
        .GetValue nama_frm, a
        index_frm.lbl_opt(a).Caption = hasilbacavalue
    Next a
    If fungsi_ato_tidak Then
        .GetValue "osh_caption", 0
        index_frm.osh.Caption = hasilbacavalue
    End If
End With
End Function

Public Sub home_set_lang()
Dim a As Integer
ini_file.FileName = App.Path & "\data\bin\xp-setup.ini"
ini_file.GetValue "Language", "lang"
lang = hasilbacavalue
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
For a = 0 To 5
    With ini_file
        .GetValue "home", a
    End With
    frm_home.Label(a).Caption = hasilbacavalue
Next a
With frm_home
    .lbl_ver(0).Caption = versi
    .lbl_ver(1).Caption = edited
    .lbl_ver(2).Caption = emailq
    .lbl_ver(3).Caption = websiteq
End With
End Sub

Public Sub about_set_lang()
With frm_about
    .Label1.Caption = "About " & versi
    .Label2.Caption = versi
    .Label4.Caption = edited
    .Label5.Caption = namaq
    .Label6.Caption = emailq
    .Label7.Caption = websiteq
End With
End Sub

Public Sub optimize_set_lang()
With frm_optimize
    If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
    If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
    .Label2.Caption = ini_file.GetValue("optimize", 0)
    '.Label3.Caption = ini_file.GetValue("optimize", 1)
    '.Label4.Caption = ini_file.GetValue("optimize", 2)
End With
End Sub

Public Function osh_punya(nama_frm As String, index_frm As Form, jmlh_fungsi As Integer) As String
Dim i As Integer
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
For i = 0 To jmlh_fungsi
    If plhn = i Then
        ini_file.GetValue nama_frm, plhn
        index_frm.osh.Caption = hasilbacavalue
    End If
Next i
End Function
