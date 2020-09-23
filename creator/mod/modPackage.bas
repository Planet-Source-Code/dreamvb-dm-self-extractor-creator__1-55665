Attribute VB_Name = "modPackage"
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

Public PackageFile As Package

Public Const def_defaut_Dir = "C:\Temp\"
Public Const def_welcome_msg = "This will now extract the files to %App_Path%"
Public Const def_finish_msg = "This program has now been installed on your computer."
Public Const def_sfx_tile = "SFX Extractor"

Public ExeHeadFile As String

Public Function DestroyPakData()
    ' clean up our package data
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
    PackageFile.RunButtonCaption = ""
End Function

Public Function WritePak(lzSaveFile As String)
' this function appends the package data to sfx\exehead.exe
Dim nFile As Long
Dim TheHeadInfo As String
    TheHeadInfo = "PAK_DATA" & Chr(5) ' Header information
    nFile = FreeFile
    Open lzSaveFile For Binary As nFile
        Put #nFile, LOF(nFile), TheHeadInfo
        Put #nFile, LOF(nFile) + 1, PackageFile
    Close #nFile
    
    lzSaveFile = ""
    
End Function

Public Function OpenFile(lzFile As String) As String
' function used to open and return a files data
Dim nFile As Long, sBuffer As String
    nFile = FreeFile
    
    Open lzFile For Binary As #nFile
        sBuffer = Space(LOF(nFile))
        Get #nFile, , sBuffer
    Close #nFile
    
    OpenFile = sBuffer
    lzFile = ""
    
End Function
