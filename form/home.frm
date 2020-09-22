VERSION 5.00
Begin VB.Form frm_home 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   300
   ClientTop       =   0
   ClientWidth     =   12735
   Icon            =   "home.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   21
      Top             =   10
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16645856
      Caption         =   "0"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_home 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Home"
      ForeColor       =   5066498
      CheckBoxMode    =   -1  'True
      Value           =   -1  'True
      HandPointer     =   -1  'True
      Picture         =   "home.frx":617A
      PictureHover    =   "home.frx":1F9EC
      pictureSize     =   3
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_visual 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Visual"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":3925E
      PictureHover    =   "home.frx":3AF68
      pictureSize     =   3
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_security 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Security"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":3CC72
      PictureHover    =   "home.frx":3E97C
      pictureSize     =   3
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_about 
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "About"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":4416E
      PictureHover    =   "home.frx":45E78
      pictureSize     =   3
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_help 
      Height          =   645
      Left            =   2880
      TabIndex        =   6
      Top             =   5400
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Help"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":47B82
      PictureHover    =   "home.frx":48A5C
      pictureSize     =   3
      MaskColor       =   16777215
   End
   Begin VB.CommandButton cmddummy 
      Caption         =   "sampah"
      Height          =   195
      Left            =   -840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8040
      Width           =   855
   End
   Begin Project1.jcbutton cmd_option 
      Height          =   645
      Left            =   5040
      TabIndex        =   7
      Top             =   5400
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Option"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":49936
      PictureHover    =   "home.frx":4B640
      pictureSize     =   3
      MaskColor       =   16777215
   End
   Begin Project1.jcbutton cmd_winfunction 
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Win Tools"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":4D34A
      PictureHover    =   "home.frx":4F054
      pictureSize     =   3
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_exit 
      Height          =   255
      Left            =   11760
      TabIndex        =   22
      Top             =   15
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16645856
      Caption         =   "x"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_optimize 
      Height          =   735
      Left            =   360
      TabIndex        =   23
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Optimize"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "home.frx":50D5E
      PictureHover    =   "home.frx":AC608
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "home.frx":107EB2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   4920
      X2              =   4920
      Y1              =   5280
      Y2              =   6120
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   615
      Index           =   5
      Left            =   7080
      TabIndex        =   20
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   2760
      Top             =   5280
      Width           =   7095
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   4
      Left            =   3000
      TabIndex        =   19
      Top             =   3840
      Width           =   6615
   End
   Begin VB.Label lbl_ver 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   3
      Left            =   10200
      TabIndex        =   18
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lbl_ver 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   17
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   975
      Index           =   3
      Left            =   10200
      TabIndex        =   16
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Readme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   10200
      TabIndex        =   14
      Top             =   3050
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   9960
      Picture         =   "home.frx":10C090
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      Height          =   3135
      Left            =   9960
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lbl_ver 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   12
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lbl_ver 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   10200
      TabIndex        =   11
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   1485
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   1490
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   9960
      Picture         =   "home.frx":10F719
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   1455
      Left            =   9960
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2760
      Picture         =   "home.frx":112DA2
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   3735
      Left            =   2760
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   1575
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   2400
      Width           =   6615
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "home.frx":11642B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    Set DockHandler.ParentForm = Me
    cek_app_requirement
    detect_lang
    cmddummy.SetFocus
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

Private Sub cmd_exit_Click()
cmddummy.SetFocus
metu
End Sub

Private Sub cmd_minimize_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub cmd_home_Click()
cmd_home.Value = True
End Sub

Private Sub cmd_visual_Click()
ganti_form frm_visual1, frm_home, Me.Left, Me.Top
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_home, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
ganti_form frm_winfunction, frm_home, Me.Left, Me.Top
End Sub

Private Sub cmd_optimize_Click()
ganti_form frm_optimize, frm_home, Me.Left, Me.Top
End Sub

Private Sub cmd_about_Click()
ganti_form frm_about, frm_home, Me.Left, Me.Top
End Sub

Private Sub cmd_option_Click()
frm_option.Show
Unload Me
End Sub

Private Sub cmd_help_Click()
On Error Resume Next
Call ShellExecute(0&, vbNullString, App.Path & "\data\bin\help.chm", vbNullString, vbNullString, vbNormalFocus)
End Sub
