VERSION 5.00
Begin VB.Form frmSfx 
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
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3585
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4965
      TabIndex        =   4
      Top             =   3585
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   165
      TabIndex        =   5
      Top             =   780
      Width           =   6075
      Begin VB.TextBox txtInformation 
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
         Height          =   2400
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   5910
      End
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "&Browse..."
      Height          =   345
      Left            =   5325
      TabIndex        =   2
      Top             =   225
      Width           =   930
   End
   Begin VB.TextBox txtInstall 
      Height          =   350
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   3975
   End
   Begin VB.Image imgLogo 
      Height          =   240
      Left            =   135
      Picture         =   "Main.frx":0000
      Top             =   3645
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Install Files to:"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    ' unload and clean up
    On Error Resume Next
    RmDir txtInstall.Text
    RemoveTemp
    CleanUp
    End
End Sub

Private Sub cmdNext_Click()
Dim Ans As Integer

   If Len(txtInstall.Text) = 3 Then
    ' ok if it three it must be a drive name
        frmSfx.Hide ' hide this form
        frmInstall.Show ' show main install form
        Exit Sub ' stop here
    End If
    
    If FindDir(txtInstall.Text) = False Then
        ' above check it the install folder is created or not
        Ans = MsgBox("The folder is not found" & vbCrLf & vbCrLf _
        & "Do you want to create the folder now?", vbYesNo Or vbQuestion)
        ' ask the user to create the folder
        If Ans = vbNo Then
            ' user selects no so we don;t do nothing simple.
            Exit Sub
        Else
            MkDir txtInstall.Text ' create the folder
            frmSfx.Hide ' hide this form
            frmInstall.Show ' show the install form
            Exit Sub ' stop
        End If
    Else
        ' folder is found so carry on
        frmSfx.Hide ' hide this form
        frmInstall.Show ' show main install form
    End If
    
End Sub

Private Sub cmdopen_Click()
Dim FolName As String
    FolName = GetFolder(frmSfx.hwnd, "Choose Folder:") ' show the user the broswe for folder dialog
    If Len(FolName) = 0 Then
        ' no folder selected so do nothing
        Exit Sub
    Else
        txtInstall.Text = FixPath(FolName) ' update the install text box with the new isntall path
        PackageFile.InstallDir = txtInstall.Text ' update the install path for the package file
    End If
    
    FolName = "" ' clear buffer
    
End Sub

Private Sub Form_Load()
    ' just loads a picture from one dialog to other ones
    frmInstall.imgLogo.Picture = imgLogo.Picture
    frmFinish.imgLogo.Picture = imgLogo.Picture
End Sub

