VERSION 5.00
Begin VB.Form frm_optimize 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   300
   ClientTop       =   -49995
   ClientWidth     =   12735
   Icon            =   "optimize.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   Begin VB.CommandButton cmddummy 
      Caption         =   "sampah"
      Height          =   195
      Left            =   -840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   855
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
      Picture         =   "optimize.frx":617A
      PictureHover    =   "optimize.frx":1F9EC
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
      Picture         =   "optimize.frx":3925E
      PictureHover    =   "optimize.frx":3AF68
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
      Picture         =   "optimize.frx":3CC72
      PictureHover    =   "optimize.frx":3E97C
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
      Picture         =   "optimize.frx":4416E
      PictureHover    =   "optimize.frx":45E78
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
      Picture         =   "optimize.frx":47B82
      PictureHover    =   "optimize.frx":4988C
      pictureSize     =   3
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   7
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
      TabIndex        =   8
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
   Begin Project1.jcbutton cmd_startup_manager 
      Height          =   735
      Left            =   2880
      TabIndex        =   9
      Top             =   1965
      Width           =   2655
      _ExtentX        =   4895
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
      Caption         =   "Startup Manager"
      ForeColor       =   5066498
      HandPointer     =   -1  'True
      Picture         =   "optimize.frx":4B596
      PictureHover    =   "optimize.frx":4C470
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_optimize 
      Height          =   735
      Left            =   360
      TabIndex        =   13
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
      CheckBoxMode    =   -1  'True
      Value           =   -1  'True
      HandPointer     =   -1  'True
      Picture         =   "optimize.frx":4D34A
      PictureHover    =   "optimize.frx":A8BF4
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   5760
      TabIndex        =   12
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   5760
      TabIndex        =   11
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   5760
      TabIndex        =   10
      Top             =   1920
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   5655
      Left            =   2760
      Top             =   1800
      Width           =   9495
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "optimize.frx":10449E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Optimize"
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
      TabIndex        =   6
      Top             =   1485
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2760
      Picture         =   "optimize.frx":10867C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   9495
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "optimize.frx":10BD05
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_optimize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    Set DockHandler.ParentForm = Me
    optimize_set_lang
    cmddummy.SetFocus
End Sub

Private Sub title_bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then DockHandler.StartDockDrag X, Y
End Sub

Private Sub title_bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then DockHandler.UpdateDockDrag X, Y
End Sub

Private Sub cmd_exit_Click()
cmddummy.SetFocus
metu
End Sub

Private Sub cmd_minimize_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub cmd_home_Click()
ganti_form frm_home, frm_optimize, Me.Left, Me.Top
End Sub

Private Sub cmd_visual_Click()
ganti_form frm_visual1, frm_optimize, Me.Left, Me.Top
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_optimize, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
ganti_form frm_winfunction, frm_optimize, Me.Left, Me.Top
End Sub

Private Sub cmd_optimize_Click()
cmd_optimize.Value = True
End Sub

Private Sub cmd_about_Click()
ganti_form frm_about, frm_optimize, Me.Left, Me.Top
End Sub

Private Sub cmd_startup_manager_Click()
ganti_form frm_startup_manager, frm_optimize, Me.Left, Me.Top
End Sub

