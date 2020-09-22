VERSION 5.00
Begin VB.Form frm_winfunction 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   -49995
   ClientWidth     =   12735
   ControlBox      =   0   'False
   Icon            =   "win_function.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   Begin Project1.jcbutton cmd_run 
      Height          =   495
      Left            =   5400
      TabIndex        =   40
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Run"
      HandPointer     =   -1  'True
   End
   Begin VB.TextBox txt_run 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   2880
      TabIndex        =   39
      Top             =   1975
      Width           =   3255
   End
   Begin VB.CommandButton cmddummy 
      Caption         =   "sampah"
      Height          =   195
      Left            =   -840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8040
      Width           =   855
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   9
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "General"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_sys 
      Height          =   495
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Log Off"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_sys 
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Restart"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_sys 
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Shutdown"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   2
      Left            =   5040
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Advanced"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Hardware"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   7
      Left            =   7560
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Regional"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   6
      Left            =   6480
      TabIndex        =   18
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Time"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   5
      Left            =   8640
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Sound"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   4
      Left            =   7560
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Themes"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   3
      Left            =   6480
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Modem"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   2
      Left            =   8640
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Internet"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Mouse"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Keyboard"
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
      HandPointer     =   -1  'True
      Picture         =   "win_function.frx":617A
      PictureHover    =   "win_function.frx":1F9EC
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
      Picture         =   "win_function.frx":3925E
      PictureHover    =   "win_function.frx":3AF68
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
      Picture         =   "win_function.frx":3CC72
      PictureHover    =   "win_function.frx":3E97C
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
      Picture         =   "win_function.frx":4416E
      PictureHover    =   "win_function.frx":45E78
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
      CheckBoxMode    =   -1  'True
      Value           =   -1  'True
      HandPointer     =   -1  'True
      Picture         =   "win_function.frx":47B82
      PictureHover    =   "win_function.frx":4988C
      pictureSize     =   3
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   1
      Left            =   10080
      TabIndex        =   21
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Group Policy Editor"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   0
      Left            =   10080
      TabIndex        =   20
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Task Manager"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   2
      Left            =   10080
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Command Prompt"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   23
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Add/Remove Programs"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   3
      Left            =   2880
      TabIndex        =   28
      Top             =   5280
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Computer Name"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   4
      Left            =   4500
      TabIndex        =   29
      Top             =   5280
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "System Restore"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   5
      Left            =   2880
      TabIndex        =   30
      Top             =   5760
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Automatic Updates"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_prop 
      Height          =   495
      Index           =   6
      Left            =   4500
      TabIndex        =   31
      Top             =   5760
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Remote"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   8
      Left            =   8640
      TabIndex        =   32
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Desktop"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   9
      Left            =   6480
      TabIndex        =   33
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Screen Saver"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   10
      Left            =   7560
      TabIndex        =   34
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Appearance"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   11
      Left            =   8640
      TabIndex        =   35
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Resolution"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   4
      Left            =   10080
      TabIndex        =   36
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "System Config. Utility"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   5
      Left            =   10080
      TabIndex        =   37
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Registry Editor"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_wintools 
      Height          =   495
      Index           =   6
      Left            =   10080
      TabIndex        =   38
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "DirectX Diagnostic Tool"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   12
      Left            =   6480
      TabIndex        =   42
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Accessibility"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   13
      Left            =   7560
      TabIndex        =   43
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "ODBC"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   14
      Left            =   8640
      TabIndex        =   44
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Infrared Port"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   15
      Left            =   6480
      TabIndex        =   45
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Joystick"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   16
      Left            =   7560
      TabIndex        =   46
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "User Manager"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   17
      Left            =   8640
      TabIndex        =   47
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Power Mngmnt"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   18
      Left            =   6480
      TabIndex        =   48
      Top             =   4800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Hardware Wizard"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_gnrl 
      Height          =   495
      Index           =   19
      Left            =   8100
      TabIndex        =   49
      Top             =   4800
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Network and Conn."
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   50
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
      TabIndex        =   51
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
   Begin Project1.jcbutton cmd_browse 
      Height          =   495
      Left            =   4560
      TabIndex        =   52
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Browse"
      HandPointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_optimize 
      Height          =   735
      Left            =   360
      TabIndex        =   53
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
      Picture         =   "win_function.frx":4B596
      PictureHover    =   "win_function.frx":A6E40
      pictureSize     =   3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "win_function.frx":1026EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alternate Run"
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
      Left            =   2880
      TabIndex        =   41
      Top             =   1485
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00404040&
      Height          =   1575
      Left            =   2760
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   2760
      Picture         =   "win_function.frx":1068C8
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Tools"
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
      Left            =   10080
      TabIndex        =   27
      Top             =   1485
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00404040&
      Height          =   3975
      Left            =   6360
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "System Properties"
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
      Left            =   2880
      TabIndex        =   25
      Top             =   4365
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "System"
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
      Left            =   2880
      TabIndex        =   24
      Top             =   3165
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "General Settings"
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
      Left            =   6480
      TabIndex        =   0
      Top             =   1485
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   9960
      Picture         =   "win_function.frx":109F51
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      Height          =   3975
      Left            =   9960
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2760
      Picture         =   "win_function.frx":10D5DA
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2760
      Picture         =   "win_function.frx":110C63
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   2055
      Left            =   2760
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   1095
      Left            =   2760
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   6360
      Picture         =   "win_function.frx":1142EC
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "win_function.frx":117975
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_winfunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    Set DockHandler.ParentForm = Me
    baca_run_history frm_winfunction
    txt_run.SetFocus
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
txt_run.SetFocus
metu
End Sub

Private Sub cmd_minimize_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub cmd_home_Click()
ganti_form frm_home, frm_winfunction, Me.Left, Me.Top
End Sub

Private Sub cmd_visual_Click()
ganti_form frm_visual1, frm_winfunction, Me.Left, Me.Top
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_winfunction, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
cmd_winfunction.Value = True
End Sub

Private Sub cmd_optimize_Click()
ganti_form frm_optimize, frm_winfunction, Me.Left, Me.Top
End Sub

Private Sub cmd_about_Click()
ganti_form frm_about, frm_winfunction, Me.Left, Me.Top
End Sub

Private Sub cmd_browse_Click()
Dim browse_file As String
browse_file = OpenFile(frm_winfunction)
If browse_file <> "" Then txt_run.Text = browse_file
End Sub

Private Sub txt_run_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then alternate_run txt_run.Text
End Sub

Private Sub cmd_run_Click()
alternate_run txt_run.Text
txt_run.SetFocus
End Sub

Private Sub cmd_sys_Click(Index As Integer)
wintools = Index
sys_logoff_restart_shutdown
End Sub

Private Sub cmd_gnrl_Click(Index As Integer)
wintools = Index
sys_general
End Sub

Private Sub cmd_prop_Click(Index As Integer)
wintools = Index
sys_prop
End Sub

Private Sub cmd_wintools_Click(Index As Integer)
wintools = Index
sys_tools
End Sub
