VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Begin VB.Form frm_about 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   300
   ClientTop       =   -49995
   ClientWidth     =   12735
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3375
      Left            =   4080
      TabIndex        =   12
      Top             =   2160
      Width           =   6615
      _cx             =   11668
      _cy             =   5953
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
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
      Picture         =   "about.frx":617A
      PictureHover    =   "about.frx":1F9EC
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
      Picture         =   "about.frx":3925E
      PictureHover    =   "about.frx":3AF68
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
      Picture         =   "about.frx":3CC72
      PictureHover    =   "about.frx":3E97C
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
      CheckBoxMode    =   -1  'True
      Value           =   -1  'True
      HandPointer     =   -1  'True
      Picture         =   "about.frx":4416E
      PictureHover    =   "about.frx":45E78
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
      Picture         =   "about.frx":47B82
      PictureHover    =   "about.frx":4988C
      pictureSize     =   3
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   13
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
      TabIndex        =   14
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
      TabIndex        =   15
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
      Picture         =   "about.frx":4B596
      PictureHover    =   "about.frx":A6E40
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "about.frx":1026EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Label Label1 
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
      Left            =   4080
      TabIndex        =   6
      Top             =   1725
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3840
      Picture         =   "about.frx":1068C8
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   5175
      Left            =   3840
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "about.frx":109F51
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    ShockwaveFlash1.Movie = App.Path & "\data\bin\about.swf"
    Set DockHandler.ParentForm = Me
    about_set_lang
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
ganti_form frm_home, frm_about, Me.Left, Me.Top
End Sub

Private Sub cmd_visual_Click()
ganti_form frm_visual1, frm_about, Me.Left, Me.Top
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_about, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
ganti_form frm_winfunction, frm_about, Me.Left, Me.Top
End Sub

Private Sub cmd_optimize_Click()
ganti_form frm_optimize, frm_about, Me.Left, Me.Top
End Sub

Private Sub cmd_about_Click()
cmd_about.Value = True
End Sub
