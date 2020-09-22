VERSION 5.00
Begin VB.Form frm_visual1 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   300
   ClientTop       =   -49995
   ClientWidth     =   12735
   Icon            =   "visual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   45
      Top             =   6720
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   44
      Top             =   6360
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   43
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   42
      Top             =   5640
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   41
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   40
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   39
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   38
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   37
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   36
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   35
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   34
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   33
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox chk_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   32
      Top             =   2040
      Width           =   255
   End
   Begin Project1.jcbutton cmd_nav_fungsi 
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   29
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "8"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_nav_fungsi 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   28
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "7"
      HandPointer     =   -1  'True
   End
   Begin VB.CommandButton cmddummy 
      Caption         =   "sampah"
      Height          =   195
      Left            =   -840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   855
   End
   Begin Project1.jcbutton apply 
      Height          =   645
      Left            =   7800
      TabIndex        =   6
      Top             =   6720
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1138
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
      Caption         =   "Apply"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      MaskColor       =   16777215
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
      HandPointer     =   -1  'True
      Picture         =   "visual.frx":617A
      PictureHover    =   "visual.frx":1F9EC
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
      CheckBoxMode    =   -1  'True
      Value           =   -1  'True
      HandPointer     =   -1  'True
      Picture         =   "visual.frx":3925E
      PictureHover    =   "visual.frx":3AF68
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
      Picture         =   "visual.frx":3CC72
      PictureHover    =   "visual.frx":3E97C
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
      Picture         =   "visual.frx":4416E
      PictureHover    =   "visual.frx":45E78
      pictureSize     =   3
      CaptionAlign    =   0
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
      Picture         =   "visual.frx":47B82
      PictureHover    =   "visual.frx":4988C
      pictureSize     =   3
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   26
      Top             =   15
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
   Begin Project1.jcbutton cmd_exit 
      Height          =   255
      Left            =   11760
      TabIndex        =   27
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
      TabIndex        =   47
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
      Picture         =   "visual.frx":4B596
      PictureHover    =   "visual.frx":A6E40
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "visual.frx":1026EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "/ 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4490
      TabIndex        =   46
      Top             =   1485
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Page:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3600
      TabIndex        =   31
      Top             =   1485
      Width           =   735
   End
   Begin VB.Label lbl_no 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      Top             =   1485
      Width           =   375
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   13
      Left            =   3360
      TabIndex        =   25
      Top             =   6720
      Width           =   4215
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   24
      Top             =   6360
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   11
      Left            =   3360
      TabIndex        =   23
      Top             =   6000
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   22
      Top             =   5640
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   21
      Top             =   5280
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   20
      Top             =   4920
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   19
      Top             =   4560
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   18
      Top             =   4200
      Width           =   6255
   End
   Begin VB.Label osh 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   2055
      Left            =   10200
      TabIndex        =   17
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   16
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   15
      Top             =   3480
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   14
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   2760
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   12
      Top             =   2400
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   11
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label lbl_opt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   1455
      Index           =   15
      Left            =   10200
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      Height          =   2295
      Left            =   9960
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label lbl_opt 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   1335
      Index           =   14
      Left            =   10200
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual"
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
      TabIndex        =   7
      Top             =   1485
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "On Screen Help"
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
      TabIndex        =   0
      Top             =   4850
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   9960
      Picture         =   "visual.frx":1068C8
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   3255
      Left            =   9960
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image img_tab_basic1 
      Height          =   375
      Left            =   2760
      Picture         =   "visual.frx":109F51
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   5655
      Left            =   2760
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   9960
      Picture         =   "visual.frx":10D5DA
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "visual.frx":110C63
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_visual1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    Set DockHandler.ParentForm = Me
    navigasi_fungsi frm_visual1, -0, False
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
ganti_form frm_home, frm_visual1, Me.Left, Me.Top
End Sub

Private Sub cmd_visual_Click()
cmd_visual.Value = True
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_visual1, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
ganti_form frm_winfunction, frm_visual1, Me.Left, Me.Top
End Sub

Private Sub cmd_optimize_Click()
ganti_form frm_optimize, frm_visual1, Me.Left, Me.Top
End Sub

Private Sub cmd_about_Click()
ganti_form frm_about, frm_visual1, Me.Left, Me.Top
End Sub

Private Sub apply_Click()
pilih_apply_fungsi frm_visual1
End Sub

Private Sub cmd_nav_fungsi_Click(Index As Integer)
navigasi_fungsi frm_visual1, Index, True
End Sub

Private Sub Chk_opt_Click(Index As Integer)
pilih_fungsi frm_visual1, Index
End Sub

Private Sub lbl_opt_Click(Index As Integer)
pilih_fungsi frm_visual1, Index
End Sub

Private Sub lbl_opt_DblClick(Index As Integer)
If Not Index = 14 And Not Index = 15 Then
    If chk_opt(Index).Value = 0 Then
        chk_opt(Index).Value = 1
    Else
        chk_opt(Index).Value = 0
    End If
End If
End Sub
