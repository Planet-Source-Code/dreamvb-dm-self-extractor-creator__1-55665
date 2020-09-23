VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM SFX Creator"
   ClientHeight    =   4230
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8010
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuild 
      Caption         =   "Build SFX"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4020
      TabIndex        =   30
      Top             =   3675
      Width           =   1110
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   645
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   1024
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   1305
      ScaleHeight     =   3360
      ScaleWidth      =   6585
      TabIndex        =   5
      Top             =   135
      Width           =   6585
      Begin VB.PictureBox stab 
         BorderStyle     =   0  'None
         Height          =   3375
         Index           =   1
         Left            =   4125
         ScaleHeight     =   3375
         ScaleWidth      =   6510
         TabIndex        =   12
         Top             =   1560
         Width           =   6510
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   315
            Left            =   930
            TabIndex        =   17
            Top             =   2955
            Width           =   1065
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   105
            TabIndex        =   16
            Top             =   2955
            Width           =   735
         End
         Begin VB.ListBox lstFiles 
            Height          =   2205
            Left            =   135
            TabIndex        =   15
            Top             =   690
            Width           =   6225
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Files:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   150
            Width           =   465
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add the files to the list you like to be installed:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   13
            Top             =   420
            Width           =   3270
         End
      End
      Begin VB.PictureBox stab 
         BorderStyle     =   0  'None
         Height          =   3375
         Index           =   2
         Left            =   2100
         ScaleHeight     =   3375
         ScaleWidth      =   6510
         TabIndex        =   18
         Top             =   345
         Width           =   6510
         Begin VB.TextBox txtRunCap 
            Height          =   315
            Left            =   135
            TabIndex        =   32
            Text            =   "Click to Run"
            Top             =   2850
            Width           =   5505
         End
         Begin VB.OptionButton opt2 
            Caption         =   "Show run button on the finish screen."
            Height          =   195
            Left            =   2175
            TabIndex        =   27
            Top             =   2505
            Width           =   3585
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Run opon user clicking the finish button."
            Height          =   195
            Left            =   2175
            TabIndex        =   26
            Top             =   2220
            Value           =   -1  'True
            Width           =   3585
         End
         Begin VB.ComboBox cboList 
            Height          =   315
            Left            =   165
            TabIndex        =   24
            Top             =   2160
            Width           =   1725
         End
         Begin VB.CheckBox chkRun 
            Caption         =   "Whould you like to run a program or command after install has finished."
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   1560
            Width           =   5610
         End
         Begin VB.TextBox txtFinish 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   705
            Width           =   6255
         End
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Button Caption"
            Height          =   195
            Left            =   150
            TabIndex        =   31
            Top             =   2565
            Width           =   1050
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "How whould you likethe program to open"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2235
            TabIndex        =   25
            Top             =   1890
            Width           =   2970
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select program or file."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   23
            Top             =   1890
            Width           =   1590
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show the message below at the end of the isntall"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   20
            Top             =   420
            Width           =   3585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finish"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   135
            TabIndex        =   19
            Top             =   150
            Width           =   510
         End
      End
      Begin VB.PictureBox stab 
         BorderStyle     =   0  'None
         Height          =   3270
         Index           =   0
         Left            =   -75
         ScaleHeight     =   3270
         ScaleWidth      =   6510
         TabIndex        =   6
         Top             =   90
         Width           =   6510
         Begin VB.TextBox txtTitle 
            Height          =   330
            Left            =   165
            TabIndex        =   29
            Top             =   1350
            Width           =   5340
         End
         Begin VB.TextBox txtWelcome 
            Height          =   990
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2100
            Width           =   5970
         End
         Begin VB.TextBox txtdefault 
            Height          =   330
            Left            =   150
            TabIndex        =   9
            Top             =   690
            Width           =   5340
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter the Self Extractors title"
            Height          =   195
            Left            =   165
            TabIndex        =   28
            Top             =   1095
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter below the information you like the user to see upon startup:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   10
            Top             =   1815
            Width           =   4710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter the default location were the files will be installed to:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   8
            Top             =   420
            Width           =   4230
         End
         Begin VB.Label lblgeneral 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "General:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   135
            TabIndex        =   7
            Top             =   150
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6750
      TabIndex        =   4
      Top             =   3660
      Width           =   1110
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   5385
      TabIndex        =   3
      Top             =   3660
      Width           =   1110
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Finish  Screen"
      Height          =   1005
      Index           =   2
      Left            =   75
      Picture         =   "frmmain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2475
      Width           =   1110
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Files"
      Height          =   1005
      Index           =   1
      Left            =   75
      Picture         =   "frmmain.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1290
      Width           =   1110
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "General Screen"
      Height          =   1005
      Index           =   0
      Left            =   75
      Picture         =   "frmmain.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   1110
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DM Self extractor creator v1
' This program allows you to package files into one exe file
' that can be then used to extract the files.
' Features in this version

' Add files
' Set the extract loaction
' Add a welcome message or other information.
' Add the extractors name eg as a title
' Add and ending message or other information
' Add open the run program when user cliks the button or presses the finish button
' Add some little tags you can use not that they were needed but I just thought make it look different

' tags are
' %App_Path%  current install path of the program
' %Date% current date
' %Time% current time
' %NoOfFiles% number of files in the self extractor

' To use them simplay place them in the finish information or welcome information text box.
' Well I hope you find this code of some use. and if you do please vote.

' Questions or answers on this code see below
' Ben Jones
' vbdream2k@yahoo.com
' or ben@eraystudios.co.uk

Dim mFilesPath As String
Private TmpStr As String
Private nRunOption As Integer
Function FindFile(lzFile As String) As Boolean
    ' check if a file is found or not
    FindFile = Dir(lzFile) <> ""
End Function

Private Sub FillComboBox()
' fill the combo box with the files the user has selected
Dim I As Long
    cboList.Clear ' clear combo box
    For I = 0 To lstFiles.ListCount - 1 ' loop until we hit the listbox's listcount
        cboList.AddItem lstFiles.List(I) ' add the items form the list box to the combo box
    Next
    I = 0 ' reset counter
    cboList.ListIndex = 0 ' default index
End Sub
Private Sub ShowTab(TabIndex As Integer)
Dim cnt As Integer
' this is used to show the picture boxes when the user clicks button on the side
    For cnt = 0 To stab.Count - 1
        stab(cnt).Visible = False ' hide all the picture box'es
    Next
    cnt = 0 ' reset counter
    
    stab(TabIndex).Left = 0 ' position the left of the selected picture box to show
    stab(TabIndex).Top = 0 ' position the top of the selected picture box to show
    stab(TabIndex).Visible = True ' show the picture box
    
End Sub

Function isFileinList(lzFile As String) As Boolean
Dim I As Long
    ' this little function is used to see if a file is already in the listbox
    isFileinList = False
    For I = 0 To lstFiles.ListCount
        If LCase(lzFile) = LCase(lstFiles.List(I)) Then
            isFileinList = True
            Exit For
        End If
    Next

End Function

Function FixPath(lzPath As String) As String
    ' fix a given path by appening a backslash if required
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Sub cboList_Change()
    cboList.Text = TmpStr
End Sub

Private Sub cboList_Click()
    TmpStr = cboList.Text
    txtRunCap.Text = "&Click to Run " & TmpStr
End Sub

Private Sub chkRun_Click()
    cboList.Enabled = chkRun
    opt1.Enabled = chkRun
    opt2.Enabled = chkRun
    txtRunCap.Enabled = chkRun
End Sub

Private Sub cmdAbout_Click()
Dim s As String
    ' just some about information
    s = "DM Self Extractor Creator" _
    & vbCrLf & vbCrLf & "By DreamVb" _
    & vbCrLf & vbCrLf & "Please vote if you like my code.."
    
    MsgBox s, vbInformation, "About..."
    s = ""
End Sub

Private Sub cmdAdd_Click()
Dim sFileNameBuff As Variant, I As Long
On Error GoTo CanErr:

    Set sFileNameBuff = Nothing
    
    With CDialog
        .CancelError = True ' turn dialog cancel error on
        .DialogTitle = "Add Files" ' setup dialog title
        .Filter = "All Files(*.*)|*.*|" ' set up the dialog filter
        .Flags = cdlOFNAllowMultiselect Or cdlOFNNoLongNames Or cdlOFNExplorer ' setup the dialog flags
        .FileName = "" ' clear the filename
        .ShowOpen ' show the open dialog
        
        If Len(.FileName) = 0 Then Exit Sub ' do nothing if no file is selected
        
        sFileNameBuff = Split(.FileName, Chr(0)) ' used to split the filenames
        mFilesPath = FixPath(CurDir) ' update the current path of the files location
        
        If UBound(sFileNameBuff) <= 0 Then ' are we lower than out lower bound count
            If isFileinList(.FileTitle) Then ' is the the already in the list
                MsgBox "The file your adding is already in the list.", vbInformation, "File Already Found"
                ' inform user that file is present in the list
                Exit Sub ' stop and so nothing
            Else
                lstFiles.AddItem .FileTitle ' add the file to the listbox
                Exit Sub ' exit do nothing
            End If
        Else
            For I = LBound(sFileNameBuff) To UBound(sFileNameBuff) ' loop though all the files in sFileNameBuff
                If isFileinList(CStr(sFileNameBuff(I + 1))) Then ' check if the file is in the list
                    MsgBox "The file your adding is already in the list.", vbInformation, "File Already Found"
                    ' yes it is so inform the user and stop
                    Exit Sub
                Else
                    ' no it ant in the list so gues what we add it then
                    lstFiles.AddItem sFileNameBuff(I + 1)
                End If
            Next
            I = 0 ' reset counter
        End If
    End With
    
    Exit Sub
CanErr:
    If Err Then
        Err.Clear
    End If
    
End Sub

Private Sub cmdBuild_Click()
Dim Counter As Long, nFile As Long
Dim AbsFilePathName As String

    If FindFile(ExeHeadFile) = False Then
        MsgBox "unable to locate the data file needed for the SFX extractor." _
        & vbCrLf & vbCrLf & "This program will now exit.", vbCritical, "File Not Found"
        ' check to see if the main sfx exe is found note needs to be in sfx\exehead.exe
        End
    End If
    
    PackageFile.SfxSig = "DM-SFX" ' header for the package file
    PackageFile.NoOfFiles = lstFiles.ListCount ' number of files in the listbox
    
    ' I don;t think I need to expain what all this is doing.
    If Len(Trim(txtdefault.Text)) = 0 Then
        PackageFile.InstallDir = def_defaut_Dir
    Else
        PackageFile.InstallDir = txtdefault.Text
    End If
    
    If Len(Trim(txtTitle.Text)) = 0 Then
        PackageFile.SfxTitle = def_sfx_tile
    Else
        PackageFile.SfxTitle = txtTitle.Text
    End If
    
    If Len(Trim(txtWelcome.Text)) = 0 Then
        PackageFile.WelcomeMsg = def_welcome_msg
    Else
        PackageFile.WelcomeMsg = txtWelcome.Text
    End If
    
    If Len(Trim(txtFinish.Text)) = 0 Then
        PackageFile.FinishMsg = def_finish_msg
    Else
        PackageFile.FinishMsg = txtFinish.Text
    End If
    '
    
    If chkRun Then ' has the user selected to run a program
        PackageFile.EnableRun = True ' yes thay have
        PackageFile.ProgramToRun = TmpStr ' add the programs name to run
        PackageFile.RunType = nRunOption ' add the users run option see finish screen
        If Len(Trim(txtRunCap.Text)) = 0 Then ' has the user entered some text for the run button
            PackageFile.RunButtonCaption = "&Click to run " & TmpStr ' no so we use the default value
        Else
            PackageFile.RunButtonCaption = "&Click to run " & txtRunCap.Text ' yes thay have so update it
        End If
    Else
        PackageFile.EnableRun = False ' nothing to run
        PackageFile.ProgramToRun = "" ' so we clear this
        PackageFile.RunType = 0 ' reset this
        PackageFile.RunButtonCaption = "" ' clear this
    End If
    
   For Counter = 0 To lstFiles.ListCount - 1 ' look thought each file in the listbox
        ReDim Preserve PackageFile.FileNames(Counter) ' resize the filename array
        ReDim Preserve PackageFile.FileData(Counter) ' resize the filedata array
        PackageFile.FileNames(Counter) = lstFiles.List(Counter) ' add in the file names to the array
        AbsFilePathName = mFilesPath & lstFiles.List(Counter) ' update the path and filename to open
        PackageFile.FileData(Counter) = OpenFile(AbsFilePathName) ' add the file data to the array form the file link above
   Next
   
   Counter = 0 ' reset counter
   AbsFilePathName = "" ' clear buffer
   
   On Error GoTo CanErr: ' error trap
   
    With CDialog
        .DialogTitle = "Save Self extractor appliaction" ' set the dialogs title
        .Filter = "Appliaction Files(*.exe)|*.exe|" ' set the dialogs file filter
        .FileName = "" ' clear the dialogs filename
        .ShowSave ' show the save dialog
        If Len(Trim(.FileName)) = 0 Then Exit Sub ' exit nothing was selected
        
        FileCopy ExeHeadFile, .FileName ' make a copy of sfx\exehead.exe
        ' and place it were the user wants to save the self extractor appliaction
        WritePak .FileName ' append the package data to the new copyed exe above
    End With
   
   Exit Sub
CanErr:
    If Err Then
        Err.Clear
    End If
    

End Sub

Private Sub cmdButton_Click(Index As Integer)

    If Index = 2 Then
        If lstFiles.ListCount = 0 Then
            MsgBox "No files have been added to the list" _
            & vbCrLf & vbCrLf & "Please add some files first.", vbInformation, "No Files Found"
            cmdBuild.Enabled = False
            Exit Sub
        End If
        
        cmdBuild.Enabled = True
        FillComboBox
    End If
    
    ShowTab Index
End Sub

Private Sub cmdexit_Click()
    Unload frmmain
End Sub

Private Sub cmdRemove_Click()
    lstFiles.RemoveItem lstFiles.ListIndex
    If lstFiles.ListIndex <= 0 Then cmdRemove.Enabled = False
End Sub

Private Sub Form_Load()
    ExeHeadFile = FixPath(App.Path) & "sfx\exehead.exe" ' main self extractor file
    txtdefault.Text = def_defaut_Dir ' update the default install loaction
    txtTitle.Text = def_sfx_tile ' update the self extractors title
    txtWelcome.Text = def_welcome_msg ' update the self extractors welcome message
    txtFinish.Text = def_finish_msg ' update the self extractors finish message
    nRunOption = 1 ' default run option set to 1
    cmdButton_Click 0
    chkRun_Click
    FindFile ExeHeadFile
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub lstFiles_Click()
    cmdRemove.Enabled = True
End Sub

Private Sub opt1_Click()
    nRunOption = 1
End Sub

Private Sub opt2_Click()
    nRunOption = 2
End Sub
