VERSION 5.00
Begin VB.Form frm_Splash 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8280
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   0
      Picture         =   "frm_Splash.frx":0000
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Load frm_main
frm_main.Show
Unload Me
End Sub
