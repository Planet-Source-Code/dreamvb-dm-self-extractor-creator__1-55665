VERSION 5.00
Begin VB.Form frmFinish 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   4950
      TabIndex        =   0
      Top             =   3585
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   165
      TabIndex        =   1
      Top             =   780
      Width           =   6075
      Begin VB.CommandButton cmdRun 
         Height          =   405
         Left            =   150
         TabIndex        =   4
         Top             =   2085
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.TextBox txtFinish 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   90
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   165
         Width           =   5895
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Install Finished"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   2
      Top             =   210
      Width           =   2640
   End
   Begin VB.Image imgLogo 
      Height          =   240
      Left            =   135
      Top             =   3645
      Width           =   1950
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmInstall.Hide
    frmSfx.Show
End Sub

Private Sub cmdCancel_Click()
    Unload frmSfx ' unload sfxform
End Sub

Private Sub cmdFinish_Click()
    If RunOnExit Then
        ' run the selected program upon exit
        RunProgram PackageFile.ProgramToRun, PackageFile.InstallDir, 1
    End If
    
    'clean up
    RemoveTemp
    CleanUp
    End
End Sub

Private Sub cmdRun_Click()
    RunProgram PackageFile.ProgramToRun, PackageFile.InstallDir, 1
    ' run the program
End Sub

