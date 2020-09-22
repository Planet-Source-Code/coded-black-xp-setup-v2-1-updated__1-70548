VERSION 5.00
Begin VB.Form frm_startup_manager 
   BorderStyle     =   0  'None
   Caption         =   "XP-Setup"
   ClientHeight    =   7995
   ClientLeft      =   300
   ClientTop       =   -49995
   ClientWidth     =   12735
   Icon            =   "startup_manager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12735
   Begin VB.Frame frame_backup 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   5890
      TabIndex        =   59
      Top             =   1825
      Visible         =   0   'False
      Width           =   6350
      Begin VB.TextBox txt_backup_name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
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
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1165
         Left            =   3360
         TabIndex        =   63
         Top             =   265
         Width           =   2775
         Begin VB.Label lbl_descrpt 
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
            Height          =   615
            Left            =   720
            TabIndex        =   68
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Dscrpt:"
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
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
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
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl_date 
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
            Height          =   375
            Left            =   720
            TabIndex        =   64
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.ListBox lst_backup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1080
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_descrpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
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
         Height          =   285
         Left            =   1320
         TabIndex        =   32
         Top             =   2280
         Width           =   4815
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   34
         Top             =   2640
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":617A
         backcolor       =   16765357
         caption         =   "Save"
         handpointer     =   -1  'True
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   33
         Top             =   2640
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":61A2
         backcolor       =   16765357
         caption         =   "Cancel"
         handpointer     =   -1  'True
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":61CA
         backcolor       =   16765357
         caption         =   "Detail"
         handpointer     =   -1  'True
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":61F2
         backcolor       =   16765357
         caption         =   "Delete"
         handpointer     =   -1  'True
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":621A
         backcolor       =   16765357
         caption         =   "Restore"
         handpointer     =   -1  'True
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Restore StartUp Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   55
         Width           =   3375
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Backup StartUp Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   1615
         Width           =   3375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   2280
         Width           =   1575
      End
   End
   Begin VB.Frame frame_backup_detail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   5890
      TabIndex        =   69
      Top             =   1825
      Visible         =   0   'False
      Width           =   6350
      Begin VB.ListBox lst_backup_detail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2340
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   6135
      End
      Begin VB.ListBox lst_backup_detail_section 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFF2&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   120
         TabIndex        =   70
         Top             =   120
         Width           =   135
      End
      Begin Project1.jcbutton cmd_backup_opt 
         Height          =   375
         Index           =   5
         Left            =   5280
         TabIndex        =   36
         Top             =   2520
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         buttonstyle     =   7
         font            =   "startup_manager.frx":6242
         backcolor       =   16765357
         caption         =   "Back"
         handpointer     =   -1  'True
      End
   End
   Begin VB.TextBox lbl_path 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
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
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   6240
      Width           =   5175
   End
   Begin VB.TextBox lbl_store 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
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
      Height          =   555
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   53
      Top             =   6600
      Width           =   5175
   End
   Begin VB.TextBox lbl_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
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
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   5520
      Width           =   5175
   End
   Begin VB.ListBox lst_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFF2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3180
      ItemData        =   "startup_manager.frx":626A
      Left            =   5880
      List            =   "startup_manager.frx":6271
      TabIndex        =   40
      Top             =   1800
      Width           =   6375
   End
   Begin VB.ComboBox Combo_store 
      BackColor       =   &H00FEFFF0&
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
      ItemData        =   "startup_manager.frx":627F
      Left            =   6720
      List            =   "startup_manager.frx":6292
      TabIndex        =   24
      Text            =   "Registry\Current User\Run"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txt_path 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFFF0&
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
      Height          =   285
      Left            =   6720
      TabIndex        =   22
      Top             =   2520
      Width           =   5415
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFFF0&
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
      Height          =   285
      Left            =   6720
      TabIndex        =   21
      Top             =   2040
      Width           =   5415
   End
   Begin Project1.jcbutton cmd_browse 
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   2880
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":6333
      backcolor       =   16765357
      caption         =   "Browse"
      handpointer     =   -1  'True
   End
   Begin VB.ListBox lst_path 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFFF0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      ItemData        =   "startup_manager.frx":635F
      Left            =   6720
      List            =   "startup_manager.frx":6366
      TabIndex        =   42
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmddummy 
      Caption         =   "sampah"
      Height          =   195
      Left            =   -840
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8040
      Width           =   855
   End
   Begin Project1.jcbutton cmd_home 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":6374
      backcolor       =   14935011
      caption         =   "Home"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":639C
      picturehover    =   "startup_manager.frx":1FC0E
      picturesize     =   3
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin Project1.jcbutton cmd_visual 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":39482
      backcolor       =   14935011
      caption         =   "Visual"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":394AA
      picturehover    =   "startup_manager.frx":3B1B4
      picturesize     =   3
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin Project1.jcbutton cmd_security 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":3CEC0
      backcolor       =   14935011
      caption         =   "Security"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":3CEE8
      picturehover    =   "startup_manager.frx":3EBF2
      picturesize     =   3
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin Project1.jcbutton cmd_about 
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4440
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":443E6
      backcolor       =   14935011
      caption         =   "About"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":4440E
      picturehover    =   "startup_manager.frx":46118
      picturesize     =   3
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin Project1.jcbutton cmd_winfunction 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":47E24
      backcolor       =   14935011
      caption         =   "Win Tools"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":47E4C
      picturehover    =   "startup_manager.frx":49B56
      picturesize     =   3
      maskcolor       =   16777215
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin Project1.jcbutton cmd_minimize 
      Height          =   255
      Left            =   11160
      TabIndex        =   38
      Top             =   15
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B862
      backcolor       =   16645856
      caption         =   "0"
      handpointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_exit 
      Height          =   255
      Left            =   11760
      TabIndex        =   39
      Top             =   15
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B88A
      backcolor       =   16645856
      caption         =   "x"
      handpointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B8AE
      backcolor       =   16765357
      caption         =   "CURRENT_USER  Run"
      handpointer     =   -1  'True
      checkboxmode    =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B8D6
      backcolor       =   16765357
      caption         =   "CURRENT_USER  Run Once"
      handpointer     =   -1  'True
      checkboxmode    =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B8FE
      backcolor       =   16765357
      caption         =   "LOCAL_MACHINE  Run"
      handpointer     =   -1  'True
      checkboxmode    =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B926
      backcolor       =   16765357
      caption         =   "LOCAL_MACHINE  Run Once"
      handpointer     =   -1  'True
      checkboxmode    =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B94E
      backcolor       =   16765357
      caption         =   "LOCAL_MACHINE  Run Services"
      handpointer     =   -1  'True
      checkboxmode    =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   11
      Top             =   3600
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B976
      backcolor       =   16765357
      caption         =   "All Users Startup Folder"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   12
      Top             =   3960
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B99E
      backcolor       =   16765357
      caption         =   "Current User Startup Folder"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   13
      Top             =   4320
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B9C6
      backcolor       =   16765357
      caption         =   "Win.ini (for advanced user only)"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupEntry 
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4B9EE
      backcolor       =   16765357
      caption         =   "System.ini (for advanced user only)"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   15
      Top             =   5520
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BA16
      backcolor       =   16765357
      caption         =   "New"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   16
      Top             =   5880
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BA3E
      backcolor       =   16765357
      caption         =   "Edit"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   17
      Top             =   6240
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BA66
      backcolor       =   16765357
      caption         =   "Delete"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_add 
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   3840
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BA8E
      backcolor       =   16765357
      caption         =   "Save"
      handpointer     =   -1  'True
   End
   Begin Project1.jcbutton cmd_cancel 
      Height          =   375
      Left            =   9960
      TabIndex        =   25
      Top             =   3840
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BAB6
      backcolor       =   16765357
      caption         =   "Cancel"
      handpointer     =   -1  'True
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFFF0&
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
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   55
      Top             =   6600
      Width           =   495
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   18
      Top             =   6600
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BADE
      backcolor       =   16765357
      caption         =   "Refresh"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   19
      Top             =   6960
      Width           =   3015
      _extentx        =   5318
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":4BB06
      backcolor       =   16765357
      caption         =   "Restore or Backup StartUp Entry"
      handpointer     =   -1  'True
      captionalign    =   0
   End
   Begin Project1.jcbutton cmd_optimize 
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
      _extentx        =   3413
      _extenty        =   1296
      buttonstyle     =   4
      font            =   "startup_manager.frx":4BB2E
      backcolor       =   14935011
      caption         =   "Optimize"
      handpointer     =   -1  'True
      picture         =   "startup_manager.frx":4BB56
      picturehover    =   "startup_manager.frx":A7400
      picturesize     =   3
      maskcolor       =   16777215
      usemaskcolor    =   -1  'True
      captionalign    =   0
      forecolor       =   5066498
   End
   Begin VB.ListBox lst_status 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEFFF0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      ItemData        =   "startup_manager.frx":102CAC
      Left            =   6720
      List            =   "startup_manager.frx":102CB3
      TabIndex        =   57
      Top             =   2040
      Width           =   735
   End
   Begin Project1.jcbutton cmd_StartupOpt 
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   20
      Top             =   5835
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      buttonstyle     =   7
      font            =   "startup_manager.frx":102CC1
      backcolor       =   16765357
      caption         =   "Ok / No"
      handpointer     =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry"
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
      Left            =   6000
      TabIndex        =   44
      Top             =   1485
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   5880
      Picture         =   "startup_manager.frx":102CE9
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Image img_status 
      Height          =   360
      Index           =   0
      Left            =   6960
      Picture         =   "startup_manager.frx":106372
      Stretch         =   -1  'True
      Top             =   5850
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image img_status 
      Height          =   360
      Index           =   1
      Left            =   6960
      Picture         =   "startup_manager.frx":109A6D
      Stretch         =   -1  'True
      Top             =   5850
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6000
      TabIndex        =   58
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lbl_status 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6960
      TabIndex        =   56
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Stored at"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6000
      TabIndex        =   51
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6000
      TabIndex        =   50
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6000
      TabIndex        =   49
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Image title_bar 
      Height          =   320
      Left            =   0
      Picture         =   "startup_manager.frx":10D2AA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Save at"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6000
      TabIndex        =   48
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   3175
      Left            =   5880
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6000
      TabIndex        =   47
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6000
      TabIndex        =   46
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   2295
      Left            =   5880
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Information"
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
      Left            =   6000
      TabIndex        =   45
      Top             =   5085
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5880
      Picture         =   "startup_manager.frx":111488
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
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
      TabIndex        =   43
      Top             =   5205
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2760
      Picture         =   "startup_manager.frx":114B11
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry On"
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
      Width           =   975
   End
   Begin VB.Image img_tab_basic1 
      Height          =   375
      Left            =   2760
      Picture         =   "startup_manager.frx":11819A
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   0
      Picture         =   "startup_manager.frx":11B823
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frm_startup_manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DockHandler As New clsDockingHandler

Private Sub Form_Load()
    Me.Show
    Set DockHandler.ParentForm = Me
    refresh_entry frm_startup_manager, Tidak
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
ganti_form frm_home, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_visual_Click()
ganti_form frm_visual1, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_security_Click()
ganti_form frm_security, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_winfunction_Click()
ganti_form frm_winfunction, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_optimize_Click()
ganti_form frm_optimize, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_about_Click()
ganti_form frm_about, frm_startup_manager, Me.Left, Me.Top
End Sub

Private Sub cmd_StartupEntry_Click(index As Integer)
pilih_startup_entry frm_startup_manager, index
End Sub

Private Sub cmd_StartupOpt_Click(index As Integer)
startup_option frm_startup_manager, index
End Sub

Private Sub Combo_store_Change()
SendKeys "{backspace}"
End Sub

Private Sub cmd_browse_Click()
Dim browse_file As String
browse_file = OpenFile(frm_startup_manager)
If browse_file <> "" Then txt_path.Text = browse_file
End Sub

Private Sub cmd_add_Click()
save_startup_entry frm_startup_manager, AddOrEdit
End Sub

Private Sub cmd_cancel_Click()
Dim ref_ As Integer
txt_name.Text = ""
txt_path.Text = ""
Combo_store.Text = "Registry\Current User\Run"
lst_name.Visible = True
For ref_ = 0 To 4
    If cmd_StartupEntry(ref_).Value = True Then refresh_entry frm_startup_manager, Semua, ref_
Next ref_
End Sub

Private Sub lst_name_Click()
startup_entry_info
End Sub

Private Sub cmd_backup_opt_Click(index As Integer)
backup_option index
End Sub

Private Sub lbl_status_Change()
Dim i As Integer
If lbl_status.Caption = "" Then
    img_status(0).Visible = False
    img_status(1).Visible = False
Else
    If lbl_status.Caption = "Enable" Then
        img_status(0).Visible = True
        img_status(1).Visible = False
    Else
        img_status(0).Visible = False
        img_status(1).Visible = True
    End If
End If
End Sub

Private Sub lst_backup_Click()
ini_file.FileName = App.Path & "\data\backup\" & lst_backup.Text & ".backup"
lbl_date.Caption = ini_file.GetValue("info", "date")
lbl_descrpt.Caption = ini_file.GetValue("info", "description")
End Sub

Private Sub lst_backup_detail_Click()
lst_backup_detail_section.ListIndex = lst_backup_detail.ListIndex
lbl_name.Text = lst_backup_detail.Text
lbl_path.Text = ini_file.GetValue(lst_backup_detail_section.Text, lst_backup_detail.Text)
lbl_store.Text = lst_backup_detail_section.Text
If lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_1 Or lbl_store.Text = "HKEY_CURRENT_USER\" & startup_entry_2 Or lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_1 Or lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_2 Or lbl_store.Text = "HKEY_LOCAL_MACHINE\" & startup_entry_3 Then
    img_status(0).Visible = True
    img_status(1).Visible = False
Else
    img_status(0).Visible = False
    img_status(1).Visible = True
End If
End Sub

