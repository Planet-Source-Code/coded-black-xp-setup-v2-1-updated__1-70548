VERSION 5.00
Begin VB.Form frm_option 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   2385
   ClientLeft      =   750
   ClientTop       =   0
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton ok 
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Ok"
      HandPointer     =   -1  'True
   End
   Begin VB.ComboBox combo_lang 
      Height          =   315
      ItemData        =   "option.frx":617A
      Left            =   1200
      List            =   "option.frx":6184
      TabIndex        =   0
      Top             =   675
      Width           =   1215
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "option.frx":619C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4725
   End
   Begin VB.Label lbl_opt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lbl_opt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   855
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lbl_opt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2370
      Left            =   0
      Picture         =   "option.frx":9AD4
      Top             =   0
      Width           =   4725
   End
End
Attribute VB_Name = "frm_option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Set DockHandler.ParentForm = Me
    ayo_load_lang = load_lang("option", frm_option, 2, False)
End Sub

Private Sub title_bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    DockHandler.StartDockDrag X, Y
End If
End Sub

Private Sub title_bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    DockHandler.UpdateDockDrag X, Y
End If
End Sub

Private Sub combo_lang_Change()
SendKeys "{backspace}"
End Sub

Private Sub ok_Click()
If lang = "En" Then ini_file.FileName = App.Path & "\data\lang\en.ini"
If lang = "Id" Then ini_file.FileName = App.Path & "\data\lang\id.ini"
With ini_file
    If combo_lang.Text = "" Then
        ini_file.GetValue "option", "1"
        MsgBox (hasilbacavalue)
        combo_lang.SetFocus
    Else
        .FileName = App.Path & "\data\bin\xp-setup.ini"
        If combo_lang.Text = "English" Then
            .WriteValue "Language", "lang", "En"
            frm_home.Show
            Unload Me
        End If
        If combo_lang.Text = "Indonesia" Then
            .WriteValue "Language", "lang", "Id"
            frm_home.Show
            Unload Me
        End If
    End If
End With
End Sub
