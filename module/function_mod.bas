Attribute VB_Name = "function_mod"
Option Explicit

Dim ngecek As String
Dim jmlh_temp As Long
Public j As Integer
Public temp As String
Public wintools As String
Public selected As String
Public temp_path As String
Public AddOrEdit As Integer
Public temp_status As String
Public ini_file As New clsINI
Public jmlh_fungsi As Integer
Public baca_perintah As String
Public open_file As OPENFILENAME
Public ini_file_temp As New clsINI
Public Reg As New RegistryFunctions

Public Sub cek_app_requirement()
On Error Resume Next
MkDir App.Path & "\data"
MkDir App.Path & "\data\backup"
MkDir App.Path & "\data\bin"
MkDir App.Path & "\data\lang"
End Sub

Public Function ngecek_setting(nm_frm As Form, fungsi_ke As Integer, tipe As String, key_kiri As String, key_kanan As String, hsl_baca As String, bener As Integer, salah As Integer, isi_default As String, Optional kosong As Boolean, Optional unique As Boolean)
With nm_frm
    Select Case tipe
    Case "dword": ngecek = Reg.GetDWordValue(key_kiri, key_kanan)
    Case "string": ngecek = Reg.GetStringValue(key_kiri, key_kanan)
    End Select
    If eon Then
        If ngecek = hsl_baca Then
            .chk_opt(fungsi_ke).Value = bener
        Else: .chk_opt(fungsi_ke).Value = salah
        End If
    Else
        If kosong Then
            ngecek = Reg.DeleteKeyValue(key_kiri, key_kanan)
            If Not unique Then
                .chk_opt(fungsi_ke).Value = bener
            Else: .chk_opt(fungsi_ke).Value = salah
            End If
        Else
            Select Case tipe
            Case "dword": ngecek = Reg.SetDWordValue(key_kiri, key_kanan, CLng(isi_default))
            Case "string": ngecek = Reg.SetStringValue(key_kiri, key_kanan, isi_default)
            End Select
            .chk_opt(fungsi_ke).Value = salah
        End If
    End If
End With
End Function

Public Function apply_fungsi(nm_frm As Form, fungsi_ke As Integer, tipe As String, key_kiri As String, key_kanan As String, aktif As String, ga_aktif As String, Optional key_kiri2 As String, Optional key_kanan2 As String, Optional aktif2 As String, Optional ga_aktif2 As String, Optional kosong As Boolean, Optional unique As Boolean)
If nm_frm.chk_opt(fungsi_ke).Value = 1 Then
    If Not unique Then
        If kosong Then
            ngecek = Reg.DeleteKeyValue(key_kiri, key_kanan)
        Else: GoTo a
        End If
    Else
a:
            Select Case tipe
            Case "dword"
                ngecek = Reg.SetDWordValue(key_kiri, key_kanan, CLng(aktif))
                If Not key_kiri2 = "" And Not key_kanan2 = "" Then ngecek = Reg.SetDWordValue(key_kiri2, key_kanan2, CLng(aktif2))
            Case "string"
                ngecek = Reg.SetStringValue(key_kiri, key_kanan, aktif)
                If Not key_kiri2 = "" And Not key_kanan2 = "" Then ngecek = Reg.SetStringValue(key_kiri2, key_kanan2, aktif2)
            End Select
    End If
Else
    Select Case tipe
    Case "dword"
        ngecek = Reg.SetDWordValue(key_kiri, key_kanan, CLng(ga_aktif))
        If Not key_kiri2 = "" And Not key_kanan2 = "" Then
            If kosong Then
                ngecek = Reg.DeleteKeyValue(key_kiri2, key_kanan2)
            Else: ngecek = Reg.SetDWordValue(key_kiri2, key_kanan2, CLng(ga_aktif2))
            End If
        End If
    Case "string"
        ngecek = Reg.SetStringValue(key_kiri, key_kanan, ga_aktif)
        If Not key_kiri2 = "" And Not key_kanan2 = "" Then
            If kosong Then
                ngecek = Reg.DeleteKeyValue(key_kiri2, key_kanan2)
            Else: ngecek = Reg.SetStringValue(key_kiri2, key_kanan2, ga_aktif2)
            End If
        End If
    End Select
End If
End Function

Public Function pilih_apply_fungsi(nm_frm As Form)
With nm_frm
    Select Case .Name
    Case "frm_visual1"
        Select Case .lbl_no.Caption
        Case 1: visual1_apply
        Case 2: visual2_apply
        End Select
    Case "frm_security"
        Select Case .lbl_no.Caption
        Case 1: security_apply
        Case 2: security2_apply
        End Select
    End Select
.cmddummy.SetFocus
End With
after_apply
End Function

Public Function navigasi_fungsi(nm_frm As Form, index As Integer, ganti_hlmn As Boolean)
Dim Key As String
With nm_frm
    If ganti_hlmn Then
        Select Case index
        Case 0: If Not .lbl_no.Caption = 1 Then .lbl_no.Caption = .lbl_no.Caption - 1
        Case 1: If Not .lbl_no.Caption = 2 Then .lbl_no.Caption = .lbl_no.Caption + 1
        End Select
    End If
    Select Case .Name
    Case "frm_visual1"
        Select Case .lbl_no.Caption
        Case 1
            visual1_load
            Key = "visual1"
        Case 2
            visual2_load
            Key = "visual2"
        End Select
    Case "frm_security"
        Select Case .lbl_no.Caption
        Case 1
            security_load
            Key = "security1"
        Case 2
            security2_load
            Key = "security2"
        End Select
    End Select
    ayo_load_lang = load_lang(Key, nm_frm, 15, True)
    fungsi_select nm_frm
    .cmddummy.SetFocus
End With
End Function

Public Function pilih_fungsi(nm_frm As Form, index As Integer)
Dim Key As String
If Not index = 14 And Not index = 15 Then
    With nm_frm
        Select Case .Name
        Case "frm_visual1"
            Select Case .lbl_no.Caption
            Case 1: Key = "osh_visual1"
            Case 2: Key = "osh_visual2"
            End Select
        Case "frm_security"
            Select Case .lbl_no.Caption
            Case 1: Key = "osh_security1"
            Case 2: Key = "osh_security2"
            End Select
        End Select
        fungsi_select_set nm_frm, index
        ayo_load_osh = osh_punya(Key, nm_frm, 13)
    End With
End If
End Function

Public Function fungsi_select(nm_frm As Form)
Dim i As Integer
With nm_frm
    For i = 0 To 13
        .lbl_opt(i).BorderStyle = 0
    Next i
End With
End Function

Public Function fungsi_select_set(nm_frm As Form, index As Integer)
Dim i As Integer
With nm_frm
For i = 0 To 13
    If .chk_opt.Item(i).index = index Then
        .lbl_opt.Item(i).BorderStyle = 1
        plhn = i
    Else: .lbl_opt.Item(i).BorderStyle = 0
    End If
Next i
End With
End Function

Public Function cek_app_instance(nm_frm As Form)
Dim m_hWnd As Long
With nm_frm
    m_hWnd = FindWindow(vbNullString, .Caption)
    If Not m_hWnd = 1 Then
        Unload nm_frm
        Exit Function
    End If
    .Show
End With
End Function

Public Function ganti_form(ganti_ke_nm_frm As Form, ganti_dari_nm_frm As Form, X As Integer, Y As Integer)
With ganti_ke_nm_frm
    .Show
    .Left = X
    .Top = Y
    Unload ganti_dari_nm_frm
End With
End Function

Public Sub visual1_load()
ngecek_setting frm_visual1, 0, "string", vis_f_0_kiri, vis_f_0_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 1, "dword", vis_f_1_kiri, vis_f_1_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 2, "dword", vis_f_2_kiri, vis_f_2_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 3, "dword", vis_f_3_kiri, vis_f_3_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 4, "string", vis_f_4_kiri, vis_f_4_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 5, "string", vis_f_5_kiri, vis_f_5_kanan, 0, 1, 0, 400
ngecek_setting frm_visual1, 6, "dword", MostUsedkey, vis_f_6_kanan, 1, 0, 1, 0
ngecek_setting frm_visual1, 7, "dword", vis_f_7_kiri, vis_f_7_kanan, 1, 0, 1, 0
ngecek_setting frm_visual1, 8, "dword", MostUsedkey, vis_f_8_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 9, "dword", vis_f_9_kiri, vis_f_9_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 10, "string", vis_f_10_kiri, vis_f_10_kanan, "Y", 0, 1, "Y"
ngecek_setting frm_visual1, 11, "dword", vis_f_11_kiri, vis_f_11_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 12, "dword", vis_f_12_kiri, vis_f_12_kanan, 0, 1, 0, 1
ngecek_setting frm_visual1, 13, "dword", vis_f_13_kiri, vis_f_13_kanan, 0, 1, 0, 1
End Sub

Public Sub visual2_load()
ngecek_setting frm_visual1, 0, "dword", vis2_f_0_kiri, vis2_f_0_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 1, "dword", vis2_f_1_kiri, vis2_f_1_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 2, "dword", vis2_f_2_kiri, vis2_f_2_kanan, 1, 0, 1, 1
ngecek_setting frm_visual1, 3, "string", vis2_f_3_kiri, vis2_f_3_kanan, "Yes", 1, 0, "No"
ngecek_setting frm_visual1, 4, "string", vis2_f_4_kiri, vis2_f_4_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 5, "dword", MostUsedkey, vis2_f_5_kanan, 95, 1, 0, 91
ngecek_setting frm_visual1, 6, "string", vis2_f_6_kiri, vis2_f_6_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 7, "dword", vis2_f_7_kiri, vis2_f_7_kanan, 1, 0, 1, 1
ngecek_setting frm_visual1, 8, "dword", vis2_f_8_kiri, vis2_f_8_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 9, "dword", vis2_f_9_kiri, vis2_f_9_kanan, 1, 0, 1, 1
ngecek_setting frm_visual1, 10, "dword", vis2_f_10_kiri, vis2_f_10_kanan, 1, 1, 0, 0
ngecek_setting frm_visual1, 11, "string", vis2_f_11_kiri, vis2_f_11_kanan, 2, 1, 0, 0
ngecek_setting frm_visual1, 12, "string", vis2_f_12_kiri, vis2_f_12_kanan, 7, 1, 0, 0
ngecek_setting frm_visual1, 13, "dword", vis2_f_13_kiri, vis2_f_13_kanan, 1, 1, 0, 0
End Sub

Public Sub security_load()
ngecek_setting frm_security, 0, "dword", sec_f_0_kiri, sec_f_0_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 1, "dword", MostUsedkey, sec_f_1_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 2, "dword", sec_f_2_kiri, sec_f_2_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 3, "dword", MostUsedkey, sec_f_3_0_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 4, "dword", MostUsedkey, sec_f_4_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 5, "dword", MostUsedkey, sec_f_5_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 6, "dword", sec_f_6_0_kiri, sec_f_6_0_kanan, 0, 1, 0, 0, True, True
ngecek_setting frm_security, 7, "dword", MostUsedkey, sec_f_7_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 8, "dword", sec_f_8_kiri, sec_f_8_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 9, "dword", sec_f_9_kiri, sec_f_9_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 10, "dword", sec_f_10_kiri, sec_f_10_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 11, "string", sec_f_11_kiri, sec_f_11_kanan, "", 1, 0, "mstask.exe", True
ngecek_setting frm_security, 12, "dword", sec_f_12_kiri, sec_f_12_kanan, 2, 1, 0, 0
ngecek_setting frm_security, 13, "dword", sec_f_13_kiri, sec_f_13_kanan, 1, 1, 0, 0
End Sub

Public Sub security2_load()
ngecek_setting frm_security, 0, "dword", MostUsedkey, sec2_f_0_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 1, "dword", MostUsedkey, sec2_f_1_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 2, "dword", MostUsedkey, sec2_f_2_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 3, "dword", MostUsedkey, sec2_f_3_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 4, "dword", MostUsedkey, sec2_f_4_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 5, "dword", MostUsedkey, sec2_f_5_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 6, "dword", MostUsedkey, sec2_f_6_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 7, "dword", MostUsedkey, sec2_f_7_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 8, "dword", MostUsedkey, sec2_f_8_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 9, "dword", sec2_f_9_kiri, sec2_f_9_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 10, "dword", sec2_f_10_kiri, sec2_f_10_kanan, 0, 1, 0, 1
ngecek_setting frm_security, 11, "dword", MostUsedkey, sec2_f_11_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 12, "dword", sec2_f_12_0_kiri, sec2_f_12_0_kanan, 1, 1, 0, 0
ngecek_setting frm_security, 13, "dword", MostUsedkey, sec2_f_13_kanan, 1, 1, 0, 0
End Sub

Public Sub visual1_apply()
apply_fungsi frm_visual1, 0, "string", vis_f_0_kiri, vis_f_0_kanan, 0, 1
apply_fungsi frm_visual1, 1, "dword", vis_f_1_kiri, vis_f_1_kanan, 0, 1
apply_fungsi frm_visual1, 2, "dword", vis_f_2_kiri, vis_f_2_kanan, 0, 1
apply_fungsi frm_visual1, 3, "dword", vis_f_3_kiri, vis_f_3_kanan, 1, 0
apply_fungsi frm_visual1, 4, "string", vis_f_4_kiri, vis_f_4_kanan, 0, 1
apply_fungsi frm_visual1, 5, "string", vis_f_5_kiri, vis_f_5_kanan, 0, 400
apply_fungsi frm_visual1, 6, "dword", MostUsedkey, vis_f_6_kanan, 0, 1
apply_fungsi frm_visual1, 7, "dword", vis_f_7_kiri, vis_f_7_kanan, 0, 1
apply_fungsi frm_visual1, 8, "dword", MostUsedkey, vis_f_8_kanan, 1, 0
apply_fungsi frm_visual1, 9, "dword", vis_f_9_kiri, vis_f_9_kanan, 0, 1
apply_fungsi frm_visual1, 10, "string", vis_f_10_kiri, vis_f_10_kanan, "N", "Y"
apply_fungsi frm_visual1, 11, "dword", vis_f_11_kiri, vis_f_11_kanan, 1, 0
apply_fungsi frm_visual1, 12, "dword", vis_f_12_kiri, vis_f_12_kanan, 0, 1
apply_fungsi frm_visual1, 13, "dword", vis_f_13_kiri, vis_f_13_kanan, 0, 1
End Sub

Public Sub visual2_apply()
apply_fungsi frm_visual1, 0, "dword", vis2_f_0_kiri, vis2_f_0_kanan, 1, 0
apply_fungsi frm_visual1, 1, "dword", vis2_f_1_kiri, vis2_f_1_kanan, 1, 0
apply_fungsi frm_visual1, 2, "dword", vis2_f_2_kiri, vis2_f_2_kanan, 0, 1
apply_fungsi frm_visual1, 3, "string", vis2_f_3_kiri, vis2_f_3_kanan, "Yes", "No"
apply_fungsi frm_visual1, 4, "string", vis2_f_4_kiri, vis2_f_4_kanan, 1, 0
apply_fungsi frm_visual1, 5, "dword", MostUsedkey, vis2_f_5_kanan, 95, 91
apply_fungsi frm_visual1, 6, "string", vis2_f_6_kiri, vis2_f_6_kanan, 1, 0
apply_fungsi frm_visual1, 7, "dword", vis2_f_7_kiri, vis2_f_7_kanan, 0, 1
apply_fungsi frm_visual1, 8, "dword", vis2_f_8_kiri, vis2_f_8_kanan, 1, 0
apply_fungsi frm_visual1, 9, "dword", vis2_f_9_kiri, vis2_f_9_kanan, 0, 1
apply_fungsi frm_visual1, 10, "dword", vis2_f_10_kiri, vis2_f_10_kanan, 1, 0
apply_fungsi frm_visual1, 11, "string", vis2_f_11_kiri, vis2_f_11_kanan, 2, 0
apply_fungsi frm_visual1, 12, "string", vis2_f_12_kiri, vis2_f_12_kanan, 7, 0
apply_fungsi frm_visual1, 13, "dword", vis2_f_13_kiri, vis2_f_13_kanan, 1, 0
End Sub

Public Sub security_apply()
apply_fungsi frm_security, 0, "dword", sec_f_0_kiri, sec_f_0_kanan, 1, 0
apply_fungsi frm_security, 1, "dword", MostUsedkey, sec_f_1_kanan, 1, 0
apply_fungsi frm_security, 2, "dword", sec_f_2_kiri, sec_f_2_kanan, 1, 0
apply_fungsi frm_security, 3, "dword", MostUsedkey, sec_f_3_0_kanan, 1, 0, sec_f_3_1_kiri, sec_f_3_1_kanan, 0, 2
apply_fungsi frm_security, 4, "dword", MostUsedkey, sec_f_4_kanan, 1, 0
apply_fungsi frm_security, 5, "dword", MostUsedkey, sec_f_5_kanan, 1, 0
apply_fungsi frm_security, 6, "dword", sec_f_6_1_kiri, sec_f_6_1_kanan, 1, 0, sec_f_6_0_kiri, sec_f_6_0_kanan, 0, , True, True
apply_fungsi frm_security, 7, "dword", MostUsedkey, sec_f_7_kanan, 1, 0
apply_fungsi frm_security, 8, "dword", sec_f_8_kiri, sec_f_8_kanan, 1, 0
apply_fungsi frm_security, 9, "dword", sec_f_9_kiri, sec_f_9_kanan, 1, 0
apply_fungsi frm_security, 10, "dword", sec_f_10_kiri, sec_f_10_kanan, 1, 0
apply_fungsi frm_security, 11, "string", sec_f_11_kiri, sec_f_11_kanan, "", "mstask.exe", , , , , True
apply_fungsi frm_security, 12, "dword", sec_f_12_kiri, sec_f_12_kanan, 2, 0
apply_fungsi frm_security, 13, "dword", sec_f_13_kiri, sec_f_13_kanan, 1, 0
End Sub

Public Sub security2_apply()
apply_fungsi frm_security, 0, "dword", MostUsedkey, sec2_f_0_kanan, 1, 0
apply_fungsi frm_security, 1, "dword", MostUsedkey, sec2_f_1_kanan, 1, 0
apply_fungsi frm_security, 2, "dword", MostUsedkey, sec2_f_2_kanan, 1, 0
apply_fungsi frm_security, 3, "dword", MostUsedkey, sec2_f_3_kanan, 1, 0
apply_fungsi frm_security, 4, "dword", MostUsedkey, sec2_f_4_kanan, 1, 0
apply_fungsi frm_security, 5, "dword", MostUsedkey, sec2_f_5_kanan, 1, 0
apply_fungsi frm_security, 6, "dword", MostUsedkey, sec2_f_6_kanan, 1, 0
apply_fungsi frm_security, 7, "dword", MostUsedkey, sec2_f_7_kanan, 1, 0
apply_fungsi frm_security, 8, "dword", MostUsedkey, sec2_f_8_kanan, 1, 0
apply_fungsi frm_security, 9, "dword", sec2_f_9_kiri, sec2_f_9_kanan, 1, 0
apply_fungsi frm_security, 10, "dword", sec2_f_10_kiri, sec2_f_10_kanan, 0, 1
apply_fungsi frm_security, 11, "dword", MostUsedkey, sec2_f_11_kanan, 1, 0
apply_fungsi frm_security, 12, "dword", sec2_f_12_0_kiri, sec2_f_12_0_kanan, 1, 0, sec2_f_12_0_kiri, sec2_f_12_1_kanan, 1, 0
apply_fungsi frm_security, 13, "dword", MostUsedkey, sec2_f_13_kanan, 1, 0
End Sub

Public Function alternate_run(perintah As String)
Dim return_ As Long
With ini_file
    If Not perintah = "" Then
        On Error Resume Next
        return_ = ShellExecute(0&, vbNullString, perintah, vbNullString, vbNullString, vbNormalFocus)
        If return_ = 42 Or return_ = 33 Then
            .FileName = App.Path & "\data\bin\xp-setup.ini"
            .WriteValue "Run_History", "last_run", perintah
        Else
            If lang = "En" Then .FileName = App.Path & "\en.ini"
            If lang = "Id" Then .FileName = App.Path & "\id.ini"
            If return_ = 2 Then MsgBox (.GetValue("Run", 0) & perintah & .GetValue("Run", 1))
        End If
    Else
        If lang = "En" Then .FileName = App.Path & "\en.ini"
        If lang = "Id" Then .FileName = App.Path & "\id.ini"
        MsgBox (.GetValue("Run", 2))
    End If
End With
End Function

Public Function baca_run_history(nm_frm As Form)
ini_file.FileName = App.Path & "\data\bin\xp-setup.ini"
With ini_file
    If .GetValue("Run_History", "last_run") = "" Then
        .WriteValue "Run_History", "last_run", ""
    Else: nm_frm.txt_run.Text = .GetValue("Run_History", "last_run")
    End If
End With
End Function

Public Sub sys_prop()
Dim perintah As String, str2exec As Double
Select Case wintools
Case 0: perintah = sys_prop_0
Case 1: perintah = sys_prop_1
Case 2: perintah = sys_prop_2
Case 3: perintah = sys_prop_3
Case 4: perintah = sys_prop_4
Case 5: perintah = sys_prop_5
Case 6: perintah = sys_prop_6
End Select
On Error Resume Next
str2exec = Shell(perintah, 5)
End Sub

Public Sub sys_general()
Dim perintah As String, str2exec As Double
Select Case wintools
Case 0: perintah = sys_gnrl_0
Case 1: perintah = sys_gnrl_1
Case 2: perintah = sys_gnrl_2
Case 3: perintah = sys_gnrl_3
Case 4: perintah = sys_gnrl_4
Case 5: perintah = sys_gnrl_5
Case 6: perintah = sys_gnrl_6
Case 7: perintah = sys_gnrl_7
Case 8: perintah = sys_gnrl_8
Case 9: perintah = sys_gnrl_9
Case 10: perintah = sys_gnrl_10
Case 11: perintah = sys_gnrl_11
Case 12: perintah = sys_gnrl_12
Case 13: perintah = sys_gnrl_13
Case 14: perintah = sys_gnrl_14
Case 15: perintah = sys_gnrl_15
Case 16: perintah = sys_gnrl_16
Case 17: perintah = sys_gnrl_17
Case 18: perintah = sys_gnrl_18
Case 19: perintah = sys_gnrl_19
End Select
On Error Resume Next
str2exec = Shell(perintah, 5)
End Sub

Public Sub sys_tools()
Dim perintah As String, str2exec As Double
Select Case wintools
Case 0: perintah = sys_tools_0
Case 1: perintah = sys_tools_1
Case 2: perintah = sys_tools_2
Case 3: perintah = sys_tools_3
Case 4: perintah = sys_tools_4
Case 5: perintah = sys_tools_5
Case 6: perintah = sys_tools_6
End Select
On Error GoTo waduh_salah
If Not wintools = 3 Then
    Call ShellExecute(0&, vbNullString, perintah, vbNullString, vbNullString, vbNormalFocus)
Else: str2exec = Shell(perintah, 5)
End If
waduh_salah:
End Sub

Public Function metu()
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
ini_file.GetValue "tutup", 0
If MsgBox(hasilbacavalue, vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
    Unload frm_home
    Unload frm_visual1
    Unload frm_security
    Unload frm_winfunction
    Unload frm_optimize
    Unload frm_about
    Unload frm_option
    Unload frm_startup_manager
    End
End If
End Function

Public Sub after_apply()
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
ini_file.GetValue "after_apply", 0
If MsgBox(hasilbacavalue, vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then Shell (logoff_)
End Sub

Public Function WinDir() As String
Dim retval As String, Tmp As String
Tmp = Space$(255)
retval = GetWindowsDirectory(Tmp, 255)
WinDir = Trim$(Left$(Tmp, retval))
End Function

Public Sub sys_logoff_restart_shutdown()
Dim perintah As String, pesan As String, i As Integer
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
ini_file.GetValue "system", wintools
Select Case wintools
Case 0
    pesan = hasilbacavalue
    perintah = shutdown_
Case 1
    pesan = hasilbacavalue
    perintah = restart_
Case 2
    pesan = hasilbacavalue
    perintah = logoff_
End Select
If MsgBox(pesan, vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then Shell (perintah)
End Sub

Function cek_lokasi_folder(Folder As Folders) As String
Dim sPath As String, retval As Long
sPath = String(MAX_PATH, 0)
retval = SHGetFolderPath(0, Folder Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, sPath)
Select Case retval
    Case S_OK
        sPath = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
        cek_lokasi_folder = sPath
    Case S_FALSE: cek_lokasi_folder = ""
    Case E_INVALIDARG: cek_lokasi_folder = ""
End Select
End Function

Public Function OpenKey(RegistryKey As RegistryKeys, Optional SubKey As String) As Long
If OpenKey <> 0 Then RegCloseKey (OpenKey)
RegOpenKeyEx RegistryKey, SubKey, 0, KEY_QUERY_VALUE, OpenKey
End Function

Public Function GetCount(RegisteryKeyHandle As Long, ValuesOrKeys As ValKey) As Long
If ValuesOrKeys = Keys Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, GetCount, 0, 0, 0, 0, 0, 0, 0
If ValuesOrKeys = Values Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, GetCount, 0, MAX_PATH + 1, 0, 0
End Function

Public Function EnumKey(RegisteryKeyHandle As Long, KeyIndex As Long) As String
EnumKey = Space(MAX_PATH + 1)
RegEnumKey RegisteryKeyHandle, KeyIndex, EnumKey, MAX_PATH + 1
EnumKey = Trim(EnumKey)
End Function

Public Function EnumValue(RegisteryKeyHandle As Long, KeyIndex As Long) As String
Dim lBufferLen As Long, i As Integer
For i = 0 To 255
    baData.ByteBuffer(i) = 0
Next
lBufferLen = 255
EnumValue = Space(MAX_PATH + 1)
RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
RegEnumValue RegisteryKeyHandle, KeyIndex, EnumValue, MAX_PATH + 1, 0, 0, baData.FirstByte, lBufferLen
EnumValue = Trim(EnumValue)
ini_file.FileName = App.Path & "\data\bin\xp-setup.ini"
ini_file.WriteValue "Temporary", "temp", EnumValue
EnumValue = ini_file.GetValue("Temporary", "temp")
ini_file.WriteValue "Temporary", "temp", ""
End Function

Public Function GetKeyValue(hKey As Long, Keyname As String) As String
Dim i As Long, rc As Long, hDepth As Long, sKeyVal As String, lKeyValType As Long, tmpVal As String, KeyValSize As Long
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(hKey, Keyname, 0, lKeyValType, tmpVal, KeyValSize)
GetKeyValue = Trim(tmpVal)
ini_file.FileName = App.Path & "\data\bin\xp-setup.ini"
ini_file.WriteValue "Temporary", "temp", GetKeyValue
GetKeyValue = ini_file.GetValue("Temporary", "temp")
ini_file.WriteValue "Temporary", "temp", ""
End Function

Public Function pilih_startup_entry(nm_frm As Form, index As Integer)
Dim i As Integer
On Error Resume Next
With nm_frm
    jmlh_temp = 0
    .Text1.Text = ""
    .lbl_name.Text = ""
    .lbl_path.Text = ""
    .lbl_store.Text = ""
    .lbl_status.Caption = ""
    .cmd_StartupOpt(5).Enabled = False
    If index <= 4 Then
        For i = 0 To 4
            If .cmd_StartupEntry(i).index = index Then
                .cmd_StartupEntry(i).Value = True
            Else: .cmd_StartupEntry(i).Value = False
            End If
        Next i
    End If
    Select Case index
    Case 0
        baca_startup_entry nm_frm, 0, HKEY_CURRENT_USER, startup_entry_1
        baca_startup_entry nm_frm, 0, HKEY_CURRENT_USER, startup_entry_1_, True
    Case 1
        baca_startup_entry nm_frm, 1, HKEY_CURRENT_USER, startup_entry_2
        baca_startup_entry nm_frm, 1, HKEY_CURRENT_USER, startup_entry_2_, True
    Case 2
        baca_startup_entry nm_frm, 2, HKEY_LOCAL_MACHINE, startup_entry_1
        baca_startup_entry nm_frm, 2, HKEY_LOCAL_MACHINE, startup_entry_1_, True
    Case 3
        baca_startup_entry nm_frm, 3, HKEY_LOCAL_MACHINE, startup_entry_2
        baca_startup_entry nm_frm, 3, HKEY_LOCAL_MACHINE, startup_entry_2_, True
    Case 4
        baca_startup_entry nm_frm, 4, HKEY_LOCAL_MACHINE, startup_entry_3
        baca_startup_entry nm_frm, 4, HKEY_LOCAL_MACHINE, startup_entry_3_, True
    Case 5: ShellExecute 0, "open", cek_lokasi_folder(Common_StartUp), "", cek_lokasi_folder(Common_StartUp), 1
    Case 6: ShellExecute 0, "open", cek_lokasi_folder(StartUp), "", cek_lokasi_folder(StartUp), 1
    Case 7: ShellExecute 0, "open", "notepad.exe", WinDir & "\win.ini", "", 1
    Case 8: ShellExecute 0, "open", "notepad.exe", WinDir & "\system.ini", "", 1
    End Select
End With
End Function

Public Function baca_startup_entry(nm_frm As Form, index As Integer, hive As RegistryKeys, Key As String, Optional nonaktif As Boolean)
Dim buka_key As Long, jmlh As Long, a As Long, b As Integer, ret_name_ As String, ret_path_ As String, cek_exist As Boolean
With nm_frm
    If Not nonaktif Then
        .lst_name.Clear
        .lst_path.Clear
        .lst_status.Clear
    End If
    buka_key = OpenKey(hive, Key)
    jmlh = GetCount(buka_key, Values)
    For a = 0 To jmlh - 1
        ret_name_ = EnumValue(buka_key, a)
        ret_path_ = GetKeyValue(buka_key, EnumValue(buka_key, a))
        If nonaktif Then
            cek_exist = False
            For b = 0 To .lst_name.ListCount - 1
                .lst_name.selected(b) = True
                .lst_path.selected(b) = True
                If .lst_name.Text = ret_name_ And .lst_path.Text = ret_path_ Then
                    cek_exist = True
                    Exit For
                End If
            Next b
            If Not cek_exist Then
                .lst_name.AddItem ret_name_
                .lst_path.AddItem ret_path_
                .lst_status.AddItem "Disable"
                jmlh = jmlh_temp + 1
                jmlh_temp = jmlh
            End If
        Else
            .lst_name.AddItem ret_name_
            .lst_path.AddItem ret_path_
            .lst_status.AddItem "Enable"
            jmlh_temp = jmlh
        End If
    Next a
    jmlh = jmlh_temp
    Select Case index
    Case 0: .cmd_StartupEntry(index).Caption = startup_capt_0 & "  (" & CStr(jmlh) & ")"
    Case 1: .cmd_StartupEntry(index).Caption = startup_capt_1 & "  (" & CStr(jmlh) & ")"
    Case 2: .cmd_StartupEntry(index).Caption = startup_capt_2 & "  (" & CStr(jmlh) & ")"
    Case 3: .cmd_StartupEntry(index).Caption = startup_capt_3 & "  (" & CStr(jmlh) & ")"
    Case 4: .cmd_StartupEntry(index).Caption = startup_capt_4 & "  (" & CStr(jmlh) & ")"
    End Select
    .lst_name.ListIndex = 0
    .lst_path.ListIndex = 0
    .lst_status.ListIndex = 0
End With
End Function

Public Function enum_startup_folder(index As Integer, folder_path As String, file_ext As String)
Dim nama_file As String, i As Long, temp As Integer, temp_0 As String
        frm_startup_manager.lst_name.Clear
        frm_startup_manager.lst_path.Clear
        frm_startup_manager.lst_status.Clear
nama_file = Dir$(folder_path & file_ext, vbArchive Or vbHidden Or vbSystem Or vbDirectory)
Do
    'DoEvents
    If Not nama_file = "" Then
        If Not (GetAttr(folder_path & nama_file) And vbDirectory) = vbDirectory And Not nama_file = "desktop.ini" Then
            frm_startup_manager.lst_name.AddItem nama_file
            frm_startup_manager.lst_path.AddItem folder_path & nama_file
            frm_startup_manager.lst_status.AddItem "Enabled"
            i = Len(nama_file)
            i = i - 7
            temp = temp + 1
        End If
    Else: Exit Do
    End If
    On Error GoTo waduh_error
    nama_file = Dir$
Loop
frm_startup_manager.cmd_StartupEntry(index).Caption = ""
Select Case index
Case 5: temp_0 = "All Users Startup Folder"
Case 6: temp_0 = "Current User Startup Folder"
End Select
frm_startup_manager.cmd_StartupEntry(index).Caption = temp_0 & "  (" & CStr(temp) & ")"
waduh_error:
End Function

Public Function startup_option(nm_frm As Form, index As Integer)
Dim value_ke As Integer, buka_key As Long, temp_0 As String, temp_1 As String, temp_2 As String, temp_3 As String, ref_ As Integer
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
On Error Resume Next
With nm_frm
    If .cmd_StartupOpt(5).Caption = "Enable" Then
        If .cmd_StartupEntry(0).Value = True Then temp_2 = "HKEY_CURRENT_USER\" & startup_entry_1
        If .cmd_StartupEntry(1).Value = True Then temp_2 = "HKEY_CURRENT_USER\" & startup_entry_2
        If .cmd_StartupEntry(2).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_1
        If .cmd_StartupEntry(3).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_2
        If .cmd_StartupEntry(4).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_3
        temp_3 = .lbl_store.Text
    Else
        If .cmd_StartupOpt(5).Caption = "Disable" Then
            If .cmd_StartupEntry(0).Value = True Then temp_2 = "HKEY_CURRENT_USER\" & startup_entry_1_
            If .cmd_StartupEntry(1).Value = True Then temp_2 = "HKEY_CURRENT_USER\" & startup_entry_2_
            If .cmd_StartupEntry(2).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_1_
            If .cmd_StartupEntry(3).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_2_
            If .cmd_StartupEntry(4).Value = True Then temp_2 = "HKEY_LOCAL_MACHINE\" & startup_entry_3_
            temp_3 = .lbl_store.Text
        End If
    End If
    temp_0 = .lbl_name.Text
    temp_1 = .lbl_path.Text
    .lbl_name.Text = ""
    .lbl_status.Caption = ""
    .lbl_path.Text = ""
    .lbl_store.Text = ""
    .cmd_StartupOpt(5).Enabled = False
    Select Case index
    Case 0
        AddOrEdit = 0
        If .cmd_StartupEntry(0).Value = True Then .Combo_store.Text = "Registry\Current User\Run"
        If .cmd_StartupEntry(1).Value = True Then .Combo_store.Text = "Registry\Current User\Run Once"
        If .cmd_StartupEntry(2).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run"
        If .cmd_StartupEntry(3).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run Once"
        If .cmd_StartupEntry(4).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run Services"
        .txt_name.SetFocus
        .lst_name.Visible = False
        startup_cmd_disable frm_startup_manager
    Case 1
        If Not .Text1.Text = "" Then
            AddOrEdit = 1
            value_ke = .lst_name.ListIndex
            .txt_name.Text = .lst_name.List(value_ke)
            .txt_path.Text = .lst_path.List(value_ke)
            If .cmd_StartupEntry(0).Value = True Then .Combo_store.Text = "Registry\Current User\Run"
            If .cmd_StartupEntry(1).Value = True Then .Combo_store.Text = "Registry\Current User\Run Once"
            If .cmd_StartupEntry(2).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run"
            If .cmd_StartupEntry(3).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run Once"
            If .cmd_StartupEntry(4).Value = True Then .Combo_store.Text = "Registry\Local Machine\Run Services"
            .txt_name.SetFocus
            .lst_name.Visible = False
            startup_cmd_disable frm_startup_manager
        Else
            ini_file.GetValue "startup", 1
            MsgBox (hasilbacavalue)
        End If
    Case 2
        If Not .Text1.Text = "" Then
            value_ke = .lst_name.ListIndex
            If MsgBox(ini_file.GetValue("startup", 3) & " '" & .lst_name.List(value_ke) & "' " & ini_file.GetValue("startup", 4), vbQuestion Or vbYesNo) = vbYes Then
                Reg.DeleteKeyValue temp_path, selected
                For ref_ = 0 To 4
                    If .cmd_StartupEntry(ref_).Value = True Then refresh_entry frm_startup_manager, Tidak, ref_
                Next ref_
            End If
        Else
            ini_file.GetValue "startup", 2
            MsgBox (hasilbacavalue)
        End If
    Case 3
        For ref_ = 0 To 4
            If .cmd_StartupEntry(ref_).Value = True Then refresh_entry frm_startup_manager, Tidak, ref_
        Next ref_
    Case 4
        enum_backup_entry frm_startup_manager, App.Path & "\data\backup\", frm_startup_manager.lst_backup
        .frame_backup.Visible = True
        .txt_backup_name.SetFocus
        startup_cmd_disable frm_startup_manager
    Case 5
        If MsgBox("Are you sure you want to " & .cmd_StartupOpt(5).Caption & " " & .lst_name.Text & "?", vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
            Reg.DeleteKeyValue temp_3, temp_0
            Reg.SetStringValue temp_2, temp_0, temp_1
            For ref_ = 0 To 4
                If .cmd_StartupEntry(ref_).Value = True Then refresh_entry frm_startup_manager, Tidak, ref_
            Next ref_
        End If
    End Select
End With
End Function
 
Public Function OpenFile(nm_frm As Form) As String
Dim hasil_ As Long
With open_file
    .lStructSize = Len(open_file)
    .hWndOwner = nm_frm.hwnd
    .hInstance = App.hInstance
    .lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*"
    .lpstrFile = Space$(254)
    .nMaxFile = 255
    .lpstrFileTitle = Space$(254)
    .nMaxFileTitle = 255
    .lpstrTitle = "Browse"
    hasil_ = GetOpenFileName(open_file)
    If (hasil_) Then OpenFile = Trim(.lpstrFile)
End With
End Function

Public Function save_startup_entry(nm_frm As Form, index As Integer)
Dim ref_ As Integer
With nm_frm
    If Not Trim(.txt_name.Text) = "" And Not Trim(.txt_path.Text) = "" And Not Trim(.Combo_store.Text) = "" Then
        Select Case index
        Case 0
            Select Case .Combo_store.Text
            Case "Registry\Current User\Run": Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_1, .txt_name.Text, .txt_path.Text
            Case "Registry\Current User\Run Once": Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_2, .txt_name.Text, .txt_path.Text
            Case "Registry\Local Machine\Run": Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_1, .txt_name.Text, .txt_path.Text
            Case "Registry\Local Machine\Run Once": Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_2, .txt_name.Text, .txt_path.Text
            Case "Registry\Local Machine\Run Services": Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_3, .txt_name.Text, .txt_path.Text
            End Select
        Case 1
            Reg.DeleteKeyValue temp_path, selected
            Select Case .Combo_store.Text
            Case "Registry\Current User\Run"
                If temp_status = "Enable" Then
                    Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_1, .txt_name.Text, .txt_path.Text
                Else: If temp_status = "Disable" Then Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_1_, .txt_name.Text, .txt_path.Text
                End If
            Case "Registry\Current User\Run Once"
                If temp_status = "Enable" Then
                    Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_2, .txt_name.Text, .txt_path.Text
                Else: If temp_status = "Disable" Then Reg.SetStringValue "HKEY_CURRENT_USER\" & startup_entry_2_, .txt_name.Text, .txt_path.Text
                End If
            Case "Registry\Local Machine\Run"
                If temp_status = "Enable" Then
                    Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_1, .txt_name.Text, .txt_path.Text
                Else: If temp_status = "Disable" Then Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_1_, .txt_name.Text, .txt_path.Text
                End If
            Case "Registry\Local Machine\Run Once"
                If temp_status = "Enable" Then
                    Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_2, .txt_name.Text, .txt_path.Text
                Else: If temp_status = "Disable" Then Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_2_, .txt_name.Text, .txt_path.Text
                End If
            Case "Registry\Local Machine\Run Services"
                If temp_status = "Enable" Then
                    Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_3, .txt_name.Text, .txt_path.Text
                Else: If temp_status = "Disable" Then Reg.SetStringValue "HKEY_LOCAL_MACHINE\" & startup_entry_3_, .txt_name.Text, .txt_path.Text
                End If
            End Select
        End Select
        .lst_name.Visible = True
        .txt_name.Text = ""
        .txt_path.Text = ""
        .Combo_store.Text = "Registry\Current User\Run"
        For ref_ = 0 To 4
            If .cmd_StartupEntry(ref_).Value = True Then refresh_entry frm_startup_manager, Semua, ref_
        Next ref_
    Else
        If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
        If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
        ini_file.GetValue "startup", 0
        MsgBox (hasilbacavalue)
    End If
End With
End Function

Public Function startup_cmd_enable(nm_frm As Form)
With nm_frm
    For j = 0 To 8
        .cmd_StartupEntry(j).Enabled = True
    Next j
    For j = 0 To 4
        .cmd_StartupOpt(j).Enabled = True
    Next j
    .cmd_home.Enabled = True
    .cmd_visual.Enabled = True
    .cmd_security.Enabled = True
    .cmd_winfunction.Enabled = True
    .cmd_optimize.Enabled = True
    .cmd_about.Enabled = True
End With
End Function

Public Function startup_cmd_disable(nm_frm As Form)
With nm_frm
    For j = 0 To 8
        .cmd_StartupEntry(j).Enabled = False
    Next j
    For j = 0 To 4
        .cmd_StartupOpt(j).Enabled = False
    Next j
    .cmd_home.Enabled = False
    .cmd_visual.Enabled = False
    .cmd_security.Enabled = False
    .cmd_winfunction.Enabled = False
    .cmd_optimize.Enabled = False
    .cmd_about.Enabled = False
End With
End Function

Public Function refresh_entry(nm_frm As Form, semua_ato_tidak As ref_semua_ato_tidak, Optional index As Integer)
With nm_frm
    If semua_ato_tidak = Semua Then GoTo a
    If semua_ato_tidak = Tidak Then GoTo b
a:
    startup_cmd_enable frm_startup_manager
b:
    j = 5
    Do Until j = 0
        j = j - 1
        pilih_startup_entry nm_frm, j
    Loop
    enum_startup_folder 5, "F:\Documents and Settings\All Users\Start Menu\Programs\Startup\", "*.*"
    enum_startup_folder 6, "F:\Documents and Settings\coded_black\Start Menu\Programs\Startup\", "*.*"
    pilih_startup_entry nm_frm, index
End With
End Function

Public Function startup_entry_info()
With frm_startup_manager
    .lst_path.selected(.lst_name.ListIndex) = True
    .lst_status.selected(.lst_name.ListIndex) = True
    .Text1.Text = .lst_name.Text & "  =  " & .lst_path.Text
    .lbl_name.Text = .lst_name.Text
    .lbl_path.Text = .lst_path.Text
    .lbl_status.Caption = .lst_status.Text
    If .lbl_status.Caption = "Enable" Then
        If .cmd_StartupEntry(0).Value = True Then .lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_1
        If .cmd_StartupEntry(1).Value = True Then .lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_2
        If .cmd_StartupEntry(2).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_1
        If .cmd_StartupEntry(3).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_2
        If .cmd_StartupEntry(4).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_3
    Else
        If .lbl_status.Caption = "Disable" Then
            If .cmd_StartupEntry(0).Value = True Then .lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_1_
            If .cmd_StartupEntry(1).Value = True Then .lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_2_
            If .cmd_StartupEntry(2).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_1_
            If .cmd_StartupEntry(3).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_2_
            If .cmd_StartupEntry(4).Value = True Then .lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_3_
        End If
    End If
    If .lbl_name.Text = "" Then
        .cmd_StartupOpt(5).Enabled = False
    Else
        .cmd_StartupOpt(5).Enabled = True
        If .lbl_status.Caption = "Enable" Then .cmd_StartupOpt(5).Caption = "Disable"
        If .lbl_status.Caption = "Disable" Then .cmd_StartupOpt(5).Caption = "Enable"
    End If
    selected = .lst_name.List(.lst_name.ListIndex)
    temp_status = .lbl_status.Caption
    temp_path = .lbl_store.Text
End With
End Function

Public Function save_startup_backup(nama As String, deskripsi As String, Optional temp_0 As String, Optional temp_1 As String, Optional temp_2 As String, Optional temp_3 As String)
Dim i As Integer, j As Integer, date_time_ As String
With frm_startup_manager
    j = 5
    date_time_ = CStr(Year(Now)) & "_" & CStr(Month(Now)) & "_" & CStr(Day(Now)) & "_" & CStr(Hour(Now)) & "_" & CStr(Minute(Now)) & "_" & CStr(Second(Now))
    Do Until j = 0
        j = j - 1
        pilih_startup_entry frm_startup_manager, j
        If Not .lst_name.ListCount = 0 Then
            For i = 1 To .lst_name.ListCount
                .lst_name.selected(i - 1) = True
                .lst_path.selected(i - 1) = True
                .lst_status.selected(i - 1) = True
                temp_0 = .lst_name.Text
                temp_1 = .lst_path.Text
                temp_2 = .lst_status.Text
                If .cmd_StartupEntry(0).Value = True Then
                    If .lst_status.Text = "Enable" Then
                        temp_3 = "HKEY_CURRENT_USER\" & startup_entry_1
                    Else: temp_3 = "HKEY_CURRENT_USER\" & startup_entry_1_
                    End If
                End If
                If .cmd_StartupEntry(1).Value = True Then
                    If .lst_status.Text = "Enable" Then
                        temp_3 = "HKEY_CURRENT_USER\" & startup_entry_2
                    Else: temp_3 = "HKEY_CURRENT_USER\" & startup_entry_2_
                    End If
                End If
                If .cmd_StartupEntry(2).Value = True Then
                    If .lst_status.Text = "Enable" Then
                        temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_1
                    Else: temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_1_
                    End If
                End If
                If .cmd_StartupEntry(3).Value = True Then
                    If .lst_status.Text = "Enable" Then
                        temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_2
                    Else: temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_2_
                    End If
                End If
                If .cmd_StartupEntry(4).Value = True Then
                    If .lst_status.Text = "Enable" Then
                        temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_3
                    Else: temp_3 = "HKEY_LOCAL_MACHINE\" & startup_entry_3_
                    End If
                End If
                ini_file.FileName = App.Path & "\data\backup\" & nama & ".backup"
                ini_file.WriteValue temp_3, temp_0, temp_1
            Next i
        End If
    Loop
    .lbl_name.Text = ""
    .img_status(0).Visible = False
    .img_status(1).Visible = False
    .cmd_StartupOpt(5).Enabled = False
    .lbl_path.Text = ""
    .lbl_store.Text = ""
End With
ini_file.WriteValue "info", "date", CStr(Now)
ini_file.WriteValue "info", "description", deskripsi
End Function

Public Function enum_backup_entry(nm_frm As Form, folder_path As String, target As ListBox)
Dim ext_file As String, nama_file As String, i As Long
target.Clear
ext_file = "*.backup"
nama_file = Dir$(folder_path & ext_file, vbArchive Or vbHidden Or vbSystem Or vbDirectory)
Do
    DoEvents
    If Not nama_file = "" Then
        If Not (GetAttr(folder_path & nama_file) And vbDirectory) = vbDirectory Then
            i = Len(nama_file)
            i = i - 7
            target.AddItem Left(nama_file, i)
        End If
    Else: Exit Do
    End If
    On Error GoTo waduh_error
    nama_file = Dir$
Loop
If Not target.ListCount = 0 Then target.ListIndex = 0
waduh_error:
End Function

Public Function backup_option(index As Integer)
Dim i As Integer, j As Integer, Section As String
With frm_startup_manager
    Select Case index
    Case 0
        .frame_backup.Visible = False
        startup_cmd_enable frm_startup_manager
        .txt_backup_name.Text = ""
        .txt_descrpt.Text = ""
    Case 1
        If Not Trim(.txt_backup_name.Text) = "" And Not Trim(.txt_descrpt.Text) = "" Then
            save_startup_backup .txt_backup_name.Text, .txt_descrpt.Text
            enum_backup_entry frm_startup_manager, App.Path & "\data\backup\", .lst_backup
            .txt_backup_name.Text = ""
            .txt_descrpt.Text = ""
        Else
            MsgBox ("Any missing entry still exist. Please fill it first!")
        End If
    Case 2
        If Not .lst_backup.ListCount = 0 Then
            .frame_backup.Visible = False
            .frame_backup_detail.Visible = True
            ini_file.FileName = App.Path & "\data\backup\" & .lst_backup.Text & ".backup"
            For j = 0 To 9
                Select Case j
                Case 0: Section = "HKEY_CURRENT_USER\" & startup_entry_1
                Case 1: Section = "HKEY_CURRENT_USER\" & startup_entry_1_
                Case 2: Section = "HKEY_CURRENT_USER\" & startup_entry_2
                Case 3: Section = "HKEY_CURRENT_USER\" & startup_entry_2_
                Case 4: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_1
                Case 5: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_1_
                Case 6: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_2
                Case 7: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_2_
                Case 8: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_3
                Case 9: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_3_
                End Select
                If Not ini_file.GetAllKeys(Section).Count = 0 Then
                    For i = 1 To ini_file.GetAllKeys(Section).Count
                        .lst_backup_detail.AddItem ini_file.GetAllKeys(Section).Item(i)
                        .lst_backup_detail_section.AddItem Section
                    Next i
                End If
            Next j
            If Not .lst_backup_detail.ListCount = 0 Then .lst_backup_detail.ListIndex = 0
        Else
            MsgBox ("There is no backup available for viewing the detail, please make at least one backup first")
        End If
    Case 3
        If Not .lst_backup.ListCount = 0 Then
            If MsgBox("Are you sure you want to delete StartUp backup: '" & .lst_backup.Text & ".backup'.", vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
                Kill App.Path & "\data\backup\" & .lst_backup.Text & ".backup"
                enum_backup_entry frm_startup_manager, App.Path & "\data\backup\", .lst_backup
                If .lst_backup.ListCount = 0 Then
                    .lbl_date.Caption = ""
                    .lbl_descrpt.Caption = ""
                End If
            End If
        Else
            MsgBox ("There is no backup available to delete")
        End If
    Case 4
        If Not .lst_backup.ListCount = 0 Then
            If MsgBox("Are you sure you want to restore this backup?", vbQuestion + vbDefaultButton2 + vbYesNo) = vbYes Then
                ini_file.FileName = App.Path & "\data\backup\" & .lst_backup.Text & ".backup"
                For j = 0 To 9
                    Select Case j
                    Case 0: Section = "HKEY_CURRENT_USER\" & startup_entry_1
                    Case 1: Section = "HKEY_CURRENT_USER\" & startup_entry_1_
                    Case 2: Section = "HKEY_CURRENT_USER\" & startup_entry_2
                    Case 3: Section = "HKEY_CURRENT_USER\" & startup_entry_2_
                    Case 4: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_1
                    Case 5: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_1_
                    Case 6: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_2
                    Case 7: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_2_
                    Case 8: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_3
                    Case 9: Section = "HKEY_LOCAL_MACHINE\" & startup_entry_3_
                    End Select
                    If Not ini_file.GetAllKeys(Section).Count = 0 Then
                        For i = 1 To ini_file.GetAllKeys(Section).Count
                            Reg.SetStringValue Section, ini_file.GetAllKeys(Section).Item(i), ini_file.GetValue(Section, ini_file.GetAllKeys(Section).Item(i))
                        Next i
                    End If
                Next j
                MsgBox ("The backup has been restored succesfully")
                .frame_backup.Visible = False
                startup_cmd_enable frm_startup_manager
                .txt_backup_name.Text = ""
                .txt_descrpt.Text = ""
                refresh_entry frm_startup_manager, Tidak, 0
            End If
        Else
            MsgBox ("There is no backup available to restore, please make at least one backup first")
        End If
    Case 5
        .frame_backup.Visible = True
        .frame_backup_detail.Visible = False
        .lst_backup_detail.Clear
        .lbl_name.Text = ""
        .lbl_path.Text = ""
        .lbl_store.Text = ""
        .img_status(0).Visible = False
        .img_status(1).Visible = False
    End Select
End With
End Function

