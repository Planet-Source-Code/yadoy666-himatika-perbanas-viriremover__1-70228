VERSION 5.00
Begin VB.Form frm_about 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   -240
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   600
      LinkTimeout     =   35
      ScaleHeight     =   2625
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003CC9FF&
         Height          =   6735
         HideSelection   =   0   'False
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "frm_about.frx":0000
         Top             =   2520
         Width           =   4335
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6315
         Left            =   0
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Himatika Perbanas ViriRemover is Freeware but without any warranty.Bugs and Virus Upload Please send to yadoy666@gmail.com"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   1320
      Picture         =   "frm_about.frx":02C8
      Top             =   4440
      Width           =   4995
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

    Timer1.Enabled = True
End Sub


Private Sub Form_Load()
Dim lReturn As Long
frm_about.Show
    Timer1.Interval = 35
    VScroll1.Max = Picture1.Height
    VScroll1.Min = 0 - Text1.Height
    VScroll1.Value = VScroll1.Max


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GotoVal = Me.Height / 2


For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE


    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo


    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2


    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE


        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
Unload Me

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()


    If VScroll1.Value >= VScroll1.Min + 20 Then
         VScroll1.Value = VScroll1.Value - 35
    Else
         VScroll1.Value = VScroll1.Max

         DoEvents
        End If

            Text1.Top = VScroll1.Value
            Text1.Visible = True

            DoEvents
        End Sub




