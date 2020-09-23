VERSION 5.00
Begin VB.Form frmInstall 
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
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   2250
      TabIndex        =   3
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Install"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4965
      TabIndex        =   1
      Top             =   3585
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   165
      TabIndex        =   2
      Top             =   780
      Width           =   6075
      Begin VB.PictureBox picBase 
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   105
         ScaleHeight     =   2145
         ScaleWidth      =   5865
         TabIndex        =   5
         Top             =   150
         Width           =   5865
         Begin VB.ListBox lstFiles 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   2175
            Left            =   -15
            TabIndex        =   6
            Top             =   -15
            Width           =   5895
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ready to Install"
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
      TabIndex        =   4
      Top             =   210
      Width           =   2700
   End
   Begin VB.Image imgLogo 
      Height          =   240
      Left            =   135
      Top             =   3645
      Width           =   1950
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdBack_Click()
    frmInstall.Hide
    frmSfx.Show
End Sub

Private Sub cmdCancel_Click()
    ' end the program and clean up any temp files and folders
    On Error Resume Next
    RmDir frmSfx.txtInstall
    RemoveTemp
    CleanUp
    End
End Sub

Private Sub cmdNext_Click()
Dim Counter As Long
Dim sBuffer As String, I As Long
    
    ' the code below is the main part that extracts all the files from the package file
    For Counter = 0 To PackageFile.NoOfFiles - 1 ' loop until we reach the end of the file array
        sBuffer = PackageFile.FileData(Counter) ' get the data of each file form the array
        WriteFile FixPath(PackageFile.InstallDir) & PackageFile.FileNames(Counter), PackageFile.FileData(Counter)
        ' line above is used to write the file and it's data
        Sleep 10 ' I just added this to make it look like the program is processing the files.
        ' also the user can see the files been added to the list
        lstFiles.AddItem PackageFile.FileNames(Counter) ' add the files to the list
        lstFiles.ListIndex = Counter ' move the list index based on the current file index
    Next
    sBuffer = "" ' clear file buffer
    I = 0 ' reset counter
    Counter = 0 'reset counter
    frmInstall.Hide ' hide this form
    frmFinish.Show ' show the finish form
    lstFiles.Clear ' clear all item from the list box
    
End Sub

