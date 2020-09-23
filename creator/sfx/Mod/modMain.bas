Attribute VB_Name = "Module1"
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Type Package
    SfxSig As String * 6
    SfxTitle As String
    NoOfFiles As Long
    InstallDir As String
    WelcomeMsg As String
    FinishMsg As String
    FileNames() As String
    FileData() As String
    EnableRun As Boolean
    ProgramToRun As String
    RunButtonCaption As String
    RunType As Integer ' 1 = run after user passed exit button 2 = run with a command button on end screen
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_NEWDIALOGSTYLE As Long = &H40
Private Const BIF_RETURNONLYFSDIRS As Long = &H1

Public PackageFile As Package
Public ExeHeadFile As String
Public RunOnExit As Boolean
Public TempFile As String

Public Function RunProgram(lzProgname As String, mDirName As String, WinMode As Integer) As Long
    ' runs a program with a given name and path
    RunProgram = ShellExecute(frmSfx.hwnd, "open", lzProgname, vbNullString, mDirName, WinMode)
End Function

Function CleanUp()
    ' clean up
    DestroyPakData
    TempFile = ""
End Function

Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        Offset = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

Function FindDir(lzPath As String) As Boolean
    ' used to find a folder
    If Not Dir(FixPath(lzPath), vbDirectory) = "." Then
        FindDir = False
        Exit Function
    Else
        FindDir = True
    End If
End Function

Function RemoveTemp()
Dim S As String
    ' this is used to delete the file left behine
    ' becuase vb will not let you delete a file as you exit I have make the code below
    ' in a batch file that works very well.
    
    TempFile = GetTempPathA & "del.bat"

    S = "@ECHO OFF" & vbCrLf
    S = S & "command.com/c" & vbCrLf
    S = S & ":Try" & vbCrLf
    S = S & "del gData.tmp" & vbCrLf
    S = S & "if exist gData.tmp goto Try" & vbCrLf
    S = S & "del del.bat" & vbCrLf
    
    Open GetTempPathA & "del.bat" For Binary As #1
        Put #1, , S
    Close #1
    S = ""
    RunProgram TempFile, GetTempPathA, 0
End Function

Private Function GetTempPathA() As String
'retuen the systems temp folder path
Dim iRet As Long
Dim sBuff As String
    sBuff = Space(256)
    iRet = GetTempPath(256, sBuff)
    GetTempPathA = Left(sBuff, iRet)
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function DestroyPakData()
    ' clear our package data
    PackageFile.SfxSig = ""
    PackageFile.SfxTitle = ""
    PackageFile.NoOfFiles = 0
    PackageFile.InstallDir = ""
    PackageFile.WelcomeMsg = ""
    PackageFile.FinishMsg = ""
    Erase PackageFile.FileNames()
    Erase PackageFile.FileData()
    PackageFile.EnableRun = False
    PackageFile.ProgramToRun = ""
    PackageFile.RunType = 0
End Function

Sub Main()
Dim nFile As Long
Dim sHeader As String * 9, sBuff1 As String, sBuff2 As String, hPos As Long

    nFile = FreeFile ' free file
    sHeader = "PAK_DATA" & Chr(5) ' header in the exe file used to find the PAK data
    
    ExeHeadFile = FixPath(App.Path) & App.EXEName & ".exe" ' the path and file name of the exe were looking at now

    Open ExeHeadFile For Binary As #nFile ' open ExeHeadFile
        sBuff1 = Space(LOF(nFile)) ' resize out buffer
        Get #nFile, , sBuff1 ' get the files data in to our buffer
    Close #1
    
    hPos = InStr(1, sBuff1, sHeader, vbBinaryCompare) ' pointer to header info
    
    If hPos <= 0 Then
        ' no header found not a lot we can do about that so just end
        MsgBox "The Extractor was unable to locate the main data stream.", vbCritical, "Invaild Stream"
        End
    End If
    
    ' now we need to extract and save the main Package file
    sBuff2 = Mid(sBuff1, hPos + 9, Len(sBuff1))
    sBuff1 = "" ' we don't need this any more
    ipos = 0 ' nor do we requite this
    
    TempFile = GetTempPathA & "gData.tmp"
    Open TempFile For Binary As #1
        Put #1, , sBuff2
    Close #1
    
    Open TempFile For Binary As #2
        Get #2, , PackageFile
    Close #1
    
    sBuff2 = "" ' clean this up we ant need it now
    
    If PackageFile.SfxSig <> "DM-SFX" Then
        MsgBox "Unable to read the sfx data", vbCritical, "Invaild Header Found"
        End
    Else
       frmSfx.txtInstall.Text = FixPath(PackageFile.InstallDir)
       PackageFile.WelcomeMsg = Replace(PackageFile.WelcomeMsg, "%App_Path%", PackageFile.InstallDir)
       PackageFile.WelcomeMsg = Replace(PackageFile.WelcomeMsg, "%Date%", Date)
       PackageFile.WelcomeMsg = Replace(PackageFile.WelcomeMsg, "%Time%", Time)
       PackageFile.WelcomeMsg = Replace(PackageFile.WelcomeMsg, "%NoOfFiles%", PackageFile.NoOfFiles)
       '
       frmSfx.txtInformation.Text = PackageFile.WelcomeMsg
       PackageFile.FinishMsg = Replace(PackageFile.FinishMsg, "%App_Path%", PackageFile.InstallDir)
       PackageFile.FinishMsg = Replace(PackageFile.FinishMsg, "%Date%", Date)
       PackageFile.FinishMsg = Replace(PackageFile.FinishMsg, "%Time%", Time)
       PackageFile.FinishMsg = Replace(PackageFile.FinishMsg, "%NoOfFiles%", PackageFile.NoOfFiles)
       frmFinish.txtFinish.Text = PackageFile.FinishMsg
       frmSfx.txtInformation.Text = PackageFile.WelcomeMsg
       frmSfx.Caption = PackageFile.SfxTitle & " - DM SFX Extractor"
       frmInstall.Caption = frmSfx.Caption
       frmFinish.Caption = frmSfx.Caption
    End If
    
    If PackageFile.RunType = 1 And PackageFile.EnableRun Then
        frmFinish.cmdRun.Visible = False
        RunOnExit = True
    Else
        frmFinish.cmdRun.Caption = PackageFile.RunButtonCaption
        frmFinish.cmdRun.Visible = True
    End If
    
    If PackageFile.EnableRun = False Then
        frmFinish.cmdRun.Visible = False
        RunOnExit = False
    End If
    
    frmSfx.Show
    ExeHeadFile = ""
End Sub

Public Sub UnloadAll()
    Unload frmInstall
    Unload frmFinish
    Unload frmSfx
    Set frmInstall = Nothing
    Set frmSfx = Nothing
    Set frmFinish = Nothing
End Sub

Public Sub WriteFile(lzFile As String, StrData As String)
Dim nFile As Long
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        Put #nFile, , StrData
    Close #nFile
    
    lzFile = ""
    StrData = ""
End Sub

