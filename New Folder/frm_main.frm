VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8670
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10740
   ControlBox      =   0   'False
   FillColor       =   &H80000007&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000007&
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame_Sys 
      BackColor       =   &H80000007&
      Caption         =   "CRC32 Generator"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6975
      Left            =   2640
      TabIndex        =   47
      Top             =   960
      Visible         =   0   'False
      Width           =   7935
      Begin Himatika_AV.XpButton cmd_Apply 
         Height          =   495
         Left            =   4200
         TabIndex        =   76
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Apply"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000007&
         Caption         =   "Registry Options"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3135
         Left            =   120
         TabIndex        =   54
         Top             =   3000
         Width           =   7695
         Begin VB.CheckBox chk_mem 
            BackColor       =   &H80000007&
            Caption         =   "Tweak Memory Acces"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox chk_download 
            BackColor       =   &H80000007&
            Caption         =   "Disable Download File From IE"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   75
            Top             =   1680
            Width           =   3375
         End
         Begin VB.CheckBox chk_FTP 
            BackColor       =   &H80000007&
            Caption         =   "FTP mode On IE"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   74
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox chk_lowdisk 
            BackColor       =   &H80000007&
            Caption         =   "No Low Disk Space Warning"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   73
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CheckBox chk_file 
            BackColor       =   &H80000007&
            Caption         =   "Hide menu file From Explorer"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   72
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox chk_os 
            BackColor       =   &H80000007&
            Caption         =   "Show Operating System File"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   71
            Top             =   600
            Width           =   3255
         End
         Begin VB.CheckBox chk_hiden 
            BackColor       =   &H80000007&
            Caption         =   "Show Hidden File and Folder"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   70
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox chk_autorun 
            BackColor       =   &H80000007&
            Caption         =   "Disable AutoRun Drive/CD"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1320
            Width           =   3255
         End
         Begin VB.CheckBox chk_recent 
            BackColor       =   &H80000007&
            Caption         =   "Hide My Recent Document"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   2760
            Width           =   3255
         End
         Begin VB.CheckBox chk_Cpanel 
            BackColor       =   &H80000007&
            Caption         =   "Hide Control Panel From Start Menu"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2400
            Width           =   4095
         End
         Begin VB.CheckBox chk_pagefile 
            BackColor       =   &H80000007&
            Caption         =   "Delete page file at shutdown"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   2040
            Width           =   3255
         End
         Begin VB.CheckBox chk_shut 
            BackColor       =   &H80000007&
            Caption         =   "Tweak ShutDown Windows"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CheckBox chk_activex 
            BackColor       =   &H80000007&
            Caption         =   "No ActiveX Install From IE"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4200
            TabIndex        =   64
            Top             =   2040
            Width           =   3255
         End
         Begin VB.CheckBox chk_unloaddll 
            BackColor       =   &H80000007&
            Caption         =   "Alway Unload DLL"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox chk_start 
            BackColor       =   &H80000007&
            Caption         =   "Tweak Start Menu Show Delay"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000007&
         Caption         =   "Windows Operation"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2415
         Left            =   4920
         TabIndex        =   49
         Top             =   360
         Width           =   2895
         Begin Himatika_AV.XpButton cmd_reg 
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Registry Editor"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin Himatika_AV.XpButton cmd_task 
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Task Manager"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin Himatika_AV.XpButton cmd_cmd 
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Command Prompt"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin Himatika_AV.XpButton cmd_clean 
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1800
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Clean Manager"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000007&
         Caption         =   "Generator"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2415
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   4695
         Begin VB.TextBox txt_result 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   59
            Top             =   1320
            Width           =   4455
         End
         Begin Himatika_AV.XpButton cmd_brogen 
            Height          =   375
            Left            =   3360
            TabIndex        =   57
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Browse"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin VB.TextBox txt_pathgen 
            BackColor       =   &H00C0C0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   480
            Width           =   3135
         End
         Begin Himatika_AV.XpButton cmd_generate 
            Height          =   375
            Left            =   840
            TabIndex        =   60
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Generate"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin Himatika_AV.XpButton cmd_clear 
            Height          =   375
            Left            =   2280
            TabIndex        =   61
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Clear"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000007&
            Caption         =   "Result "
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000007&
            Caption         =   "File to Generate"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1695
         End
      End
      Begin Himatika_AV.XpButton cmd_Default 
         Height          =   495
         Left            =   5880
         TabIndex        =   77
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Default"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_cekall 
         Height          =   495
         Left            =   840
         TabIndex        =   78
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Check All"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_uncheck 
         Height          =   495
         Left            =   2520
         TabIndex        =   79
         Top             =   6360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Uncheck All"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   360
      Top             =   360
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      Caption         =   "Other Menu"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   2415
      Begin Himatika_AV.XpButton XpButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_restart 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Restart Computer"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_turnoff 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Turn Off Computer"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_about 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "About"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   2415
      Begin Himatika_AV.XpButton cmd_tweak 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CRC32 Generator"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_virscan 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Virus Scanner"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_Proc 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Process Viewer"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
   End
   Begin VB.Frame Frame_Scan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Virus Scanner"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   6975
      Left            =   2640
      TabIndex        =   12
      Top             =   960
      Width           =   7935
      Begin VB.Frame Frame_proc 
         BackColor       =   &H80000007&
         Caption         =   "Process Viewer"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   6975
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   7935
         Begin VB.Frame Frame1 
            BackColor       =   &H80000007&
            Caption         =   "Memory Information"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   2895
            Left            =   120
            TabIndex        =   29
            Top             =   3840
            Width           =   7695
            Begin VB.Shape Shape8 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   5280
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   7
               Left            =   5400
               TabIndex        =   46
               Top             =   1320
               Width           =   2145
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Computer Name"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   5280
               TabIndex        =   45
               Top             =   1080
               Width           =   1365
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   6
               Left            =   5400
               TabIndex        =   44
               Top             =   720
               Width           =   2145
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   5
               Left            =   3240
               TabIndex        =   43
               Top             =   2280
               Width           =   1785
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   2280
               Width           =   1935
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   3
               Left            =   3240
               TabIndex        =   42
               Top             =   1560
               Width           =   1665
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   1920
               Width           =   1935
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   4
               Left            =   3240
               TabIndex        =   41
               Top             =   1920
               Width           =   1785
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   1560
               Width           =   1935
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   2
               Left            =   3240
               TabIndex        =   40
               Top             =   1200
               Width           =   1785
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   1
               Left            =   3240
               TabIndex        =   39
               Top             =   840
               Width           =   1785
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   840
               Width           =   1935
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   180
               Index           =   0
               Left            =   3240
               TabIndex        =   38
               Top             =   480
               Width           =   1785
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3120
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   5400
               TabIndex        =   37
               Top             =   600
               Width           =   105
            End
            Begin VB.Shape Shape7 
               BorderColor     =   &H80000005&
               FillColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   5280
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Memory Load"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   5280
               TabIndex        =   36
               Top             =   480
               Width           =   1155
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Available Page File"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   35
               Top             =   2280
               Width           =   1995
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Total Page File"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   34
               Top             =   1920
               Width           =   1575
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Available Virtual Memory"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   33
               Top             =   1560
               Width           =   2520
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Total Virtual Memory"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   32
               Top             =   1200
               Width           =   2100
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Available Physical Memory"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   31
               Top             =   840
               Width           =   2625
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000012&
               Caption         =   "Total Physical Memory"
               BeginProperty Font 
                  Name            =   "Lucida Sans Typewriter"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   180
               Left            =   240
               TabIndex        =   30
               Top             =   480
               Width           =   2205
            End
         End
         Begin Himatika_AV.XpButton cmd_refresh 
            Height          =   495
            Left            =   4200
            TabIndex        =   28
            Top             =   3240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Refresh"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin Himatika_AV.XpButton cmd_terminate 
            Height          =   495
            Left            =   2280
            TabIndex        =   27
            Top             =   3240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Terminate"
            ForeColor       =   -2147483630
            ForeHover       =   0
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   6840
            Top             =   480
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
         End
         Begin MSComctlLib.ListView list_proc 
            Height          =   2415
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   4260
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process Name"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process Path"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "PID"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "CRC32"
               Object.Width           =   2293
            EndProperty
         End
         Begin VB.Label NUMPROC 
            AutoSize        =   -1  'True
            BackColor       =   &H80000007&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   1920
            TabIndex        =   26
            Top             =   2880
            Width           =   105
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            Caption         =   "Jumlah Process :"
            BeginProperty Font 
               Name            =   "Lucida Sans Typewriter"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   2880
            Width           =   1680
         End
      End
      Begin Himatika_AV.XpButton cmd_stop 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   5760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stop"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_delsel 
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   5760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Delete Selected"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_del 
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   5760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Delete All"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_scan 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   5760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Scan"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Himatika_AV.XpButton cmd_browse 
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Browse"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.TextBox txt_path 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   6135
      End
      Begin MSComctlLib.ListView list_viri 
         Height          =   3495
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Path/File Name"
            Object.Width           =   10231
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Virus Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CRC32"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label shpfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[-]"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   7200
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label numvir 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2400
         TabIndex        =   21
         Top             =   5160
         Width           =   225
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   "Total Virus Detected :"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "Virus Found :"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "Location To Scan :"
         BeginProperty Font 
            Name            =   "Lucida Sans Typewriter"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image3 
      Height          =   1890
      Left            =   360
      Picture         =   "frm_main.frx":08CA
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Copyright (C) 2007 YaDoY SoFtWaRe DeVeLoPmEnT. All Right Reserved"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   8400
      Width           =   7095
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   5520
      Picture         =   "frm_main.frx":CB7C
      Top             =   360
      Width           =   4995
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   -120
      Picture         =   "frm_main.frx":16416
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10980
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim WithEvents SCANPROC As cScanProcesses
Attribute SCANPROC.VB_VarHelpID = -1
Dim WithEvents SCANDIR As cScanDirectories
Attribute SCANDIR.VB_VarHelpID = -1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private t As Single

Dim CRC32 As cCRC32
Dim m_Time As Single
Dim FILEICON As cGetFileIcon

Private Sub cmd_about_Click()
frm_about.Show

End Sub

Private Sub cmd_Apply_Click()
If chk_start.Value = 0 And chk_unloaddll.Value = 0 And chk_mem.Value = 0 And chk_autorun.Value = 0 And chk_shut.Value = 0 And chk_pagefile.Value = 0 _
And chk_Cpanel.Value = 0 And chk_recent.Value = 0 And chk_hiden.Value = 0 And chk_os.Value = 0 And chk_file.Value = 0 And chk_lowdisk.Value = 0 _
And chk_FTP.Value = 0 And chk_download.Value = 0 And chk_activex.Value = 0 Then
MsgBox "Anda belum memilih Modifikasi Registry yang akan dilakukan", vbInformation + vbOKOnly, "Himatika Perbanas ViriRemover"
Else:
Call modif_reg
End If
End Sub

Private Sub cmd_brogen_Click()
 CommonDialog1.CancelError = True 'error bila user mengklik cancel pada CommonDialog
 CommonDialog1.DialogTitle = "Baca File" 'Caption commondialog
 On Error Resume Next
 CommonDialog1.ShowOpen
 txt_pathgen.Text = CommonDialog1.FileName
 

End Sub

Private Sub cmd_browse_Click()
Dim FolderPath As String
    
    FolderPath = BrowseForFolder(Me.HWND, "Pilih Drive atau Folder yang akan di scan :")
    If FolderPath <> vbNullString Then
        txt_path.Text = FolderPath
        
    End If


End Sub

Private Sub cmd_cekall_Click()
chk_start.Value = 1
chk_unloaddll.Value = 1
chk_autorun.Value = 1
chk_shut.Value = 1
chk_mem.Value = 1
chk_pagefile.Value = 1
chk_Cpanel.Value = 1
chk_recent.Value = 1
chk_hiden.Value = 1
chk_os.Value = 1
chk_file.Value = 1
chk_lowdisk.Value = 1
chk_FTP.Value = 1
chk_download.Value = 1
chk_activex.Value = 1

End Sub

Private Sub cmd_clean_Click()
On Error Resume Next
Shell GetSystemDirectory & "cleanmgr", vbNormalFocus
End Sub

Private Sub cmd_clear_Click()
txt_result.Text = ""
txt_pathgen.Text = ""
End Sub

Private Sub cmd_cmd_Click()
On Error Resume Next

Shell GetSystemDirectory & "cmd.exe", vbNormalFocus
End Sub




Private Sub cmd_Default_Click()
If MsgBox("Apakah anda yakin ingin mengembalikan Registry ke kondisi Default ?", vbInformation + vbYesNo, "Himatika Perbanas ViriRemover") = vbYes Then
Call Reg_default
MsgBox "Registry telah kembali ke settingan Default...!!", vbInformation + vbOKOnly, "Himatika Perbanas ViriRemover"
Else: Call cmd_uncheck_Click
End If

End Sub

Private Sub cmd_del_Click()
Dim i     As Integer


    If list_viri.ListItems.Count > 0 Then
        If MsgBox("Apakah anda yakin ingin menghapus semua file yang terdeteksi?", vbExclamation + vbYesNo, "Delete All") = vbYes Then
            For i = list_viri.ListItems.Count To 1 Step -1
               On Error GoTo delError
                With list_viri
                    SetAttr .ListItems(i), vbNormal
                    Kill .ListItems(i)
                    .ListItems.Remove .ListItems(i).Index
                End With
            Next i
            list_viri.SetFocus
        End If
    End If

Exit Sub

delError:
    MsgBox "Gagal menghapus file..!" & vbNewLine & _
       vbNewLine & _
       "File name : " & list_viri.ListItems(i), vbCritical, "Penghapusan Error"

Call CleanReg

End Sub

Private Sub cmd_delsel_Click()
Dim i     As Integer

    If list_viri.ListItems.Count > 0 Then
        For i = list_viri.ListItems.Count To 1 Step -1
            If list_viri.ListItems(i).Selected Then
                On Error GoTo delError
                With list_viri
                    SetAttr .SelectedItem, vbNormal
                    Kill .SelectedItem
                    .ListItems.Remove .ListItems(i).Index
                End With
            End If
        Next i
        list_viri.SetFocus
    End If
Exit Sub

delError:
    MsgBox "Gagal menghapus file yang dipilih!" & vbNewLine & _
       vbNewLine & _
       "File name : " & list_viri.ListItems(i), vbCritical, "Penghapusan Error"

Call CleanReg

End Sub



Private Sub cmd_generate_Click()
If txt_pathgen.Text = "" Then
MsgBox "Anda belum memilih file yang akan di cari CRC-nya", vbInformation + vbOKOnly, "Himatika Perbanas ViriRemover"
End If
txt_result.Text = CRC32.FileChecksum(txt_pathgen.Text)
End Sub

Private Sub cmd_Proc_Click()
Frame_proc.Visible = True
Frame_Sys.Visible = False
End Sub

Private Sub cmd_refresh_Click()
Call segarkan
End Sub

Private Sub cmd_reg_Click()
On Error Resume Next

Shell GetWindowsDirectory & "regedit.exe", vbNormalFocus

End Sub

Private Sub cmd_restart_Click()
If IsWinNT Then
        RebootNT True
    Else
        Reboot
    End If

End Sub

Private Sub cmd_scan_Click()
shpfile.Visible = True

If txt_path.Text = "" Then
MsgBox "Anda belum memasukan direktori yang akan di scan", vbInformation + vbOKOnly, "Himatika Perbanas ViriRemover"
Exit Sub
End If

cmd_scan.Enabled = False
cmd_stop.Enabled = True
cmd_del.Enabled = False
cmd_delsel.Enabled = False

list_viri.ListItems.Clear
t = Timer

        SCANDIR.StartPath = txt_path.Text
        SCANDIR.Filter = "*.exe|*.com|*.bat|*.VBS|*.pif|*.txt|*.ini|*.htt|*.inf|*.scr"
        SCANDIR.SubDirectories = True
        SCANDIR.ScanDeep = 50
        SCANDIR.BeginScanning
        
End Sub

Private Sub cmd_stop_Click()
shpfile.Visible = False
SCANDIR.CancelScanning
cmd_scan.Enabled = True
cmd_stop.Enabled = False
cmd_del.Enabled = True
cmd_delsel.Enabled = True

End Sub


Private Sub cmd_task_Click()
On Error Resume Next
Shell GetSystemDirectory & "taskmgr.exe", vbNormalFocus
End Sub

Private Sub cmd_terminate_Click()
    If MsgBox("Apakah anda yakin ingin men-terminate proses ini", vbExclamation + vbYesNo + vbDefaultButton2, "Terminate Process") = vbYes Then
        If SCANPROC.TerminateProcess(list_proc.SelectedItem.SubItems(2)) = True Then
            Call segarkan
        End If
    End If
End Sub

Private Sub cmd_turnoff_Click()
If IsWinNT Then
        ShutdownNT True
    Else
        Shutdown
    End If
End Sub

Private Sub cmd_tweak_Click()
Frame_Sys.Visible = True
End Sub


Private Sub cmd_uncheck_Click()
chk_start.Value = 0
chk_unloaddll.Value = 0
chk_autorun.Value = 0
chk_shut.Value = 0
chk_mem.Value = 0
chk_pagefile.Value = 0
chk_Cpanel.Value = 0
chk_recent.Value = 0
chk_hiden.Value = 0
chk_os.Value = 0
chk_file.Value = 0
chk_lowdisk.Value = 0
chk_FTP.Value = 0
chk_download.Value = 0
chk_activex.Value = 0
End Sub

Private Sub cmd_virscan_Click()
Frame_Scan.Visible = True
Frame_proc.Visible = False
Frame_Sys.Visible = False
txt_path.Text = ""
list_viri.ListItems.Clear
cmd_del.Enabled = False
cmd_delsel.Enabled = False
End Sub

Private Sub SCANPROC_CurrentProcess(File As String, Path As String, ID As Long, Terminate As Boolean)
    Dim p_HasImage As Boolean
    If Path <> "SYSTEM" Then
        p_HasImage = True
        On Error Resume Next
         ImageList1.ListImages(File).Tag = ""
         If Not Err.Number = 0 Then
            ImageList1.ListImages.Add , File, FILEICON.Icon(Path & File, SmallIcon)
            Err.Clear
        End If
    End If
    
    Dim lsv As ListItem
    If p_HasImage Then
        Set lsv = list_proc.ListItems.Add(, , File, , File)
    Else
        Set lsv = list_proc.ListItems.Add(, , File)
    End If
    Dim Crc As New cCRC32
    lsv.SubItems(1) = Path
    lsv.SubItems(2) = ID
    lsv.SubItems(3) = Crc.FileChecksum(Path & File)
    lsv.ListSubItems(2).ForeColor = vbBlue
    lsv.Selected = True
    lsv.EnsureVisible

If cariDatabase(Crc.FileChecksum(Path & File), "DB.yadoy") Then  'bila fungsi bernilai TRUE
    SCANPROC.TerminateProcess (list_proc.SelectedItem.SubItems(2))
    Call segarkan
    Call segarkan
    Call segarkan
End If

End Sub

Private Sub SCANPROC_DoneScanning(TotalProcesses As Integer)
    Dim p_Elapsed As Single
    p_Elapsed = Timer - m_Time
    Debug.Print "Total Number of Process Detected: " & TotalProcesses & vbNewLine & "Total Scan Time: " & p_Elapsed & vbNewLine
    NUMPROC = TotalProcesses
End Sub

Private Sub segarkan()
    list_proc.ListItems.Clear
    m_Time = Timer
    SCANPROC.BeginScanning
End Sub


Private Sub Timer1_Timer()
Dim MS     As MEMORYSTATUS
'On Run Time puts title in caption

Dim compname As String * 255, cname As String
Dim X As Long
    
    X = GetComputerName(compname, 255)
    cname = Trim(compname)
    cname = Left(cname, Len(cname) - 1)
lbl(7).Caption = cname


MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
    lbl(0) = Format$(.dwTotalPhys / 1024, "#,###.##") & "KB"
    lbl(1) = Format$(.dwAvailPhys / 1024, "#,###.##") & "KB"
    lbl(2) = Format$(.dwTotalVirtual / 1024, "#,###.##") & "KB"
    lbl(3) = Format$(.dwAvailVirtual / 1024, "#,###.##") & "KB"
    lbl(4) = Format$(.dwTotalPageFile / 1024, "#,###.##") & "KB"
    lbl(5) = Format$(.dwAvailPageFile / 1024, "#,###.##") & "KB"
    lbl(6) = Format$(.dwMemoryLoad, "##0.00") & " %"
End With

numvir.Caption = list_viri.ListItems.Count

End Sub

Private Sub XpButton1_Click()
End
End Sub

Private Sub Animate()

    If shpfile = "[-]" Then
        shpfile = "[\]"
    ElseIf shpfile = "[\]" Then
        shpfile = "[|]"
    ElseIf shpfile = "[|]" Then
        shpfile = "[/]"
    ElseIf shpfile = "[/]" Then
        shpfile = "[-]"
    End If

End Sub

Private Sub SCANDIR_CurrentFile(File As String, Path As String, Delete As Boolean)
Animate

    If Mid$(Path, 3, 2) = "\\" Then
        Path = Replace$(Path, "\\", "\")
    End If

    If cariDatabase(CRC32.FileChecksum(Path & "\" & File), "DB.yadoy") Then
    
    Dim lsv As ListItem
    If p_HasImage Then
        Set lsv = list_viri.ListItems.Add(, , File, , File)
    Else
        Set lsv = list_viri.ListItems.Add(, , File)
    End If
    
    lsv = Path & "\" & File
    lsv.SubItems(1) = namaVirus
    lsv.SubItems(2) = CrcVirus
    list_viri.ListItems(list_viri.ListItems.Count).Selected = True
    list_viri.SelectedItem.EnsureVisible
    list_viri.ForeColor = vbRed
    
    Animate
    End If
End Sub

Private Sub SCANDIR_DoneScanning(TotalFolders As Long, TotalFiles As Long)
cmd_scan.Enabled = True
cmd_stop.Enabled = False

    NUMFILES = TotalFiles
    NUMFOLDERS = TotalFolders
    numvir = list_viri.ListItems.Count
If numvir.Caption <> "0" Then
cmd_del.Enabled = True
cmd_delsel.Enabled = True
Else
cmd_del.Enabled = False
cmd_delsel.Enabled = False

End If

shpfile.Visible = False

    MsgBox "Total Folder   : " & TotalFolders & vbNewLine & _
           "Total File        : " & TotalFiles & vbNewLine & _
           "Total Virus yang terdeteksi : " & numvir & vbNewLine & vbNewLine & _
           "Total Waktu Scanning : " & Timer - t & " seconds.", vbInformation, "Scanning Selesai"
End Sub

Private Sub Form_Load()
    Set SCANDIR = New cScanDirectories
    Set CRC32 = New cCRC32
    Set SCANPROC = New cScanProcesses
    Set FILEICON = New cGetFileIcon

    Call segarkan
    list_viri.ForeColor = vbRed
    Frame_Scan.Visible = True
    Frame_proc.Visible = False
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SCANDIR.CancelScanning
    SCANPROC.CancelScanning
    
    Set SCANPROC = Nothing
    Set FILEICON = Nothing
    Set SCANDIR = Nothing
    Set CRC32 = Nothing
    
End Sub

Private Sub Form_Initialize()
    Set SCANPROC = New cScanProcesses
    Set FILEICON = New cGetFileIcon
End Sub

Private Sub modif_reg()
On Error Resume Next

If chk_start.Value = 1 Then
CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
CreateStringValue HKEY_USERS, "S-1-5-18\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
CreateStringValue HKEY_USERS, "S-1-5-19\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
CreateStringValue HKEY_USERS, "S-1-5-20\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
CreateStringValue HKEY_USERS, "S-1-5-21-602162358-261478967-682003330-1003\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "0"
End If

If chk_unloaddll.Value = 1 Then
CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\", "AlwaysUnloadDLL", 1
End If

If chk_mem.Value = 1 Then
CreateDwordValue HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "DisablePagingExecutive", 1
CreateDwordValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "LargeSystemCache", 1
End If

If chk_autorun.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun", 95
End If

If chk_shut.Value = 1 Then
CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\", REG_SZ, "WaitToKillServiceTimeout", 0
End If

If chk_pagefile.Value = 1 Then
CreateDwordValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "ClearPageFileAtShutdown", 1
End If

If chk_Cpanel.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Start_ShowControlPanel ", 1
End If

If chk_recent.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoRecentDocsMenu", 1
End If

If chk_hiden.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden", 1
End If

If chk_os.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "ShowSuperHidden", 1
End If

If chk_file.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFileMenu", 1
CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFileMenu", 1
End If

If chk_lowdisk.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoLowDiskSpaceChecks", 1
End If

If chk_FTP.Value = 1 Then
CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\FTP", REG_SZ, "Use Web Based FTP", "YES"
End If

If chk_download.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1803", 3
CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1803", 3
End If

If chk_activex.Value = 1 Then
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1001", 3
End If

MsgBox "                    Modifikasi Registry berhasil...!! " & vbNewLine & "Silahkan restart komputer anda untuk merasakan efeknya", vbInformation + vbOKOnly, "Himatika Perbanas ViriRemover"
Call cmd_uncheck_Click
End Sub
